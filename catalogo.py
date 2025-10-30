import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
# IMPORTAÇÃO NECESSÁRIA PARA O REPORTLAB LER IMAGENS
from reportlab.lib.utils import ImageReader 
from PIL import Image
from io import BytesIO
import os
import re # Importa a biblioteca de expressões regulares para normalização

# === CONFIGURAÇÕES GERAIS ===
excel_path = "produtos.xlsx"
pdf_path = "catalogo_bikeline.pdf"
logo_path = "logo_bikeline.png"  # coloque o logo da Bikeline aqui (opcional)

# NOVA CONFIGURAÇÃO: Diretório onde as imagens estão salvas
# Assumindo que a pasta "img_produtos" está no mesmo nível do script.
img_dir = "img_produtos" 

# === FUNÇÃO DE NORMALIZAÇÃO DE CÓDIGO ===
def normalize_code(code_str):
    """
    Tenta normalizar a string do código do produto para um formato que possa
    corresponder a um nome de arquivo (removendo caracteres que podem ser
    interpretados como separadores decimais).
    """
    # Garante que é uma string
    code_str = str(code_str)
    
    # 1. Tenta limpar o código removendo pontos e vírgulas (Ex: '78.0003' -> '780003')
    cleaned_code = re.sub(r'[.,]', '', code_str)
    
    # Retorna uma lista de códigos para tentar (original e limpo)
    # Inclui a versão com underscore (ex: 78_0003) caso o arquivo use isso
    return list(set([
        code_str,               # Formato original (Ex: '78.0003')
        cleaned_code,            # Formato limpo (Ex: '780003')
        code_str.replace('.', '_') # Formato com underline (Ex: '78_0003')
    ]))

# Lê a planilha
try:
    df = pd.read_excel(excel_path)
except FileNotFoundError:
    print(f"ERRO: Arquivo Excel não encontrado em: {excel_path}")
    exit()
except Exception as e:
    print(f"ERRO: Falha ao ler o arquivo Excel: {e}")
    exit()

# Cria o PDF
c = canvas.Canvas(pdf_path, pagesize=A4)
largura, altura = A4

# === FUNÇÃO PARA CABEÇALHO ===
def cabecalho(pagina):
    # Fundo de topo
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, altura - 3 * cm, largura, 3 * cm, fill=True, stroke=0)

    # Logotipo (opcional)
    try:
        # Nota: O ReportLab precisa de um caminho de arquivo real ou objeto PIL
        c.drawImage(logo_path, 2 * cm, altura - 2.7 * cm, width=3.5 * cm, preserveAspectRatio=True, mask='auto')
    except Exception as e:
        # Não para o script se o logo falhar, apenas avisa
        print(f"Aviso: Não foi possível carregar o logo de '{logo_path}'. {e}")
        pass

    # Título
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(6.5 * cm, altura - 2 * cm, "Catálogo Bikeline")

    # Linha divisória
    c.setStrokeColorRGB(0.7, 0.7, 0.7)
    c.setLineWidth(1)
    c.line(1.5 * cm, altura - 3.1 * cm, largura - 1.5 * cm, altura - 3.1 * cm)

# === FUNÇÃO PARA RODAPÉ ===
def rodape(pagina):
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(0, 0, largura, 2 * cm, fill=True, stroke=0)
    c.setFillColorRGB(0.2, 0.2, 0.2)
    c.setFont("Helvetica", 9)
    c.drawString(2 * cm, 0.8 * cm, "www.bikeline.com.br")
    c.drawRightString(largura - 2 * cm, 0.8 * cm, f"Página {pagina}")

# === CONFIGURAÇÕES DO LAYOUT ===
produtos_por_linha = 3
largura_produto = (largura - 4 * cm) / produtos_por_linha
# Aumentando a altura total do bloco do produto para acomodar a imagem maior
altura_produto = 10.5 * cm # Altura ajustada
espacamento_vertical = altura_produto + 1 * cm

# Posição inicial
pagina = 1
cabecalho(pagina)
y = altura - 4 * cm
x_inicio = 2 * cm

print("Iniciando geração do catálogo...")
erros_imagem = 0

# === LOOP DE PRODUTOS ===
for i, row in df.iterrows():
    col = i % produtos_por_linha
    x = x_inicio + col * (largura_produto + 0.8 * cm)

    # Assumindo que a coluna 'Código do Produto' é a chave para o nome do arquivo de imagem
    codigo_produto_bruto = row.get("Código do Produto", None)
    nome = str(row["Nome do Produto"])
    descricao = str(row["Descrição"])
    preco = str(row["Preço"])
    
    # 1. Bloco de fundo (desenhado primeiro para ficar atrás do texto/imagem)
    c.setFillColorRGB(0.97, 0.97, 0.97)
    # A caixa de fundo deve começar um pouco acima e terminar um pouco abaixo do bloco (y e y - altura_produto + 0.3 * cm)
    c.roundRect(x - 0.2 * cm, y - altura_produto + 0.3 * cm, largura_produto + 0.3 * cm, altura_produto - 0.3 * cm, 0.3 * cm, fill=True, stroke=0)

    # --- LÓGICA DE CARREGAMENTO DE IMAGEM LOCAL COM FLEXIBILIDADE DE CÓDIGO ---
    image_loaded = False
    caminho_imagem_encontrada = None
    
    # Altura máxima permitida para o espaço da imagem (7.0 cm)
    max_h_cm = 150
    # Topo da área reservada para a imagem
    y_area_topo = y 
    # Base da área reservada para a imagem
    y_area_base = y - max_h_cm 
    
    if codigo_produto_bruto is not None and str(codigo_produto_bruto).lower() != "nan":
        # 1. Normaliza o código bruto para tentar diferentes formatos de nome de arquivo
        codigos_tentativa = normalize_code(codigo_produto_bruto)
        # 2. Tenta extensões comuns
        extensoes = ['.jpg', '.jpeg', '.png']
        
        # 3. Tenta encontrar o arquivo em disco
        for codigo in codigos_tentativa:
            if not codigo: continue # Pula códigos vazios gerados pela normalização
            for ext in extensoes:
                caminho_tentativa = os.path.join(img_dir, f"{codigo}{ext}")
                
                if os.path.exists(caminho_tentativa):
                    caminho_imagem_encontrada = caminho_tentativa
                    image_loaded = True
                    break # Sai do loop de extensões
            if image_loaded:
                break # Sai do loop de códigos
        
        # 4. Se a imagem foi encontrada, tenta desenhar
        if image_loaded:
            try:
                # Usa ImageReader do ReportLab. Ele é mais robusto para arquivos no disco.
                img_reader = ImageReader(caminho_imagem_encontrada)
                
                # Obtém as dimensões originais da imagem
                largura_original, altura_original = img_reader.getSize()
                proporcao = largura_original / altura_original
                
                # CÁLCULO DE TAMANHO para maximizar a área dentro do box (largura_produto x max_h_cm)
                
                # Inicialmente, ajustamos pela largura máxima (preferencial)
                largura_desenho = largura_produto
                altura_desenho = largura_desenho / proporcao 
                
                # Se a altura calculada for maior que o máximo (7cm), ajusta para a altura máxima
                if altura_desenho > max_h_cm:
                    altura_desenho = max_h_cm
                    largura_desenho = altura_desenho * proporcao # Recalcula a largura para manter proporção
                
                # Posição X ajustada para CENTRALIZAR a imagem dentro do bloco (horizontalmente)
                x_ajustado = x + (largura_produto - largura_desenho) / 2
                
                # Posição Y ajustada para CENTRALIZAR a imagem dentro da área reservada (verticalmente)
                # y_area_base é a base do espaço reservado
                # (max_h_cm - altura_desenho) / 2 é o deslocamento para centralizar a imagem verticalmente
                y_ajustado = y_area_base + (max_h_cm - altura_desenho) / 2
                
                # Desenha a imagem usando o ImageReader
                c.drawImage(
                    img_reader, 
                    x_ajustado, # Usa a posição X centralizada
                    y_ajustado, # Usa a posição Y centralizada (base da imagem)
                    width=largura_desenho, 
                    height=altura_desenho, 
                    preserveAspectRatio=True
                )
                
            except Exception as e:
                erros_imagem += 1
                print(f"ERRO CRÍTICO ao desenhar a imagem '{caminho_imagem_encontrada}'. Verifique o formato do arquivo (deve ser JPG ou PNG): {e}")
                image_loaded = False # Marca como falha para desenhar placeholder
            
    # Se a imagem falhou ao carregar ou o código estava ausente
    if not image_loaded:
        erros_imagem += 1
        # Desenha o placeholder centralizado verticalmente na metade do max_h_cm
        placeholder_y = y - max_h_cm/2 - 0.2 * cm # Centralizado no espaço de 7cm
        c.setFont("Helvetica-Oblique", 10)
        c.drawString(x, placeholder_y, "Imagem não disponível") 
        if caminho_imagem_encontrada is None and str(codigo_produto_bruto).lower() != "nan":
            print(f"AVISO: Imagem local não encontrada para o código '{codigo_produto_bruto}' na pasta '{img_dir}'.")
        
    # === INSERÇÃO DO CÓDIGO DO PRODUTO ===
    # O código será posicionado 0.3cm ABAIXO da área reservada de 7cm.
    # A área termina em y - max_h_cm. Adicionamos um espaçamento (gap) de 0.5 cm.
    gap = 0.5 * cm
    codigo_y_pos = y_area_base - gap
    
    c.setFillColor(colors.darkgrey) # Cor mais discreta
    c.setFont("Helvetica", 9)
    # Garante que o código é uma string antes de desenhar
    c.drawString(x, codigo_y_pos, f"Código: {str(codigo_produto_bruto)}")


    # Nome (AJUSTADO: Posição 0.5cm abaixo do código)
    nome_y_pos = codigo_y_pos - 0.5 * cm 
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x, nome_y_pos, nome[:40]) # Limita o nome

    # Descrição (AJUSTADA: Mantendo o espaçamento relativo ao nome)
    desc_y_pos = nome_y_pos - 0.7 * cm 
    
    c.setFont("Helvetica", 9)
    desc_limpa = " ".join(descricao.split()) # Remove quebras de linha extras
    linhas_desc = [desc_limpa[i:i+40] for i in range(0, len(desc_limpa), 40)]
    for j, linha in enumerate(linhas_desc[:2]): # Pega no máximo 2 linhas
        c.drawString(x, desc_y_pos - j * 0.4 * cm, linha) 
        
    # Preço (AJUSTADO: Posicionado para ficar mais próximo da base da caixa)
    # Base da caixa é y - altura_produto + 0.3 * cm
    preco_y_pos = y - altura_produto + 0.8 * cm # 0.5cm acima da base da caixa
    
    c.setFont("Helvetica-Bold", 11)
    c.setFillColorRGB(0, 0.5, 0)
    c.drawString(x, preco_y_pos, f"R$ {preco}") 
    c.setFillColorRGB(0, 0, 0)

    # Próximo produto
    if col == produtos_por_linha - 1:
        y -= espacamento_vertical

    # Se ultrapassar o fim da página
    # A verificação deve garantir que a altura total do bloco (y - altura_produto + 0.3 * cm) ainda é maior que 2cm
    if (y - altura_produto) < 2 * cm and i + 1 < len(df): 
        rodape(pagina)
        c.showPage()
        pagina += 1
        cabecalho(pagina)
        y = altura - 4 * cm

# Rodapé final
rodape(pagina)
c.save()

print("\n--- Geração Concluída ---")
print(f"✅ Catálogo gerado com sucesso: {pdf_path}")
print(f"Total de produtos: {len(df)}")
print(f"Página{ 's' if pagina > 1 else '' } criada{ 's' if pagina > 1 else '' }: {pagina}")
if erros_imagem > 0:
    print(f"⚠️ Atenção: {erros_imagem} imagens falharam ou não foram encontradas. Verifique os 'AVISO' e 'ERRO' no console.")
