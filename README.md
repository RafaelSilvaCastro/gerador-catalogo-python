📚 Gerador de Catálogo PDF - Bike Friday A+Ciclo

Este projeto Python utiliza as bibliotecas pandas e ReportLab para gerar um catálogo de produtos profissional e responsivo em formato PDF (A4), com foco na organização alfabética dos itens. É ideal para criar documentos de vendas ou informativos de forma automatizada, utilizando dados de uma planilha Excel.

🌟 Funcionalidades Principais

Entrada de Dados Simplificada: Carrega informações diretamente de um arquivo produtos.xlsx.

Design Profissional: Inclui capa personalizada com destaque promocional em vermelho, cabeçalho e rodapé em todas as páginas.

Layout Modular de Produtos: Exibe cada produto em um "card" com:

Imagem do produto.

Código (destacado em vermelho forte).

Descrição (fonte e espaçamento reduzidos para concisão).

Preços (preço antigo riscado e preço promocional em vermelho).

Otimização de Layout: A descrição do produto é ajustada dinamicamente para caber no espaço limitado do card.

Ordenação Fixa: Atualmente configurado para listar todos os produtos em ordem alfabética pela descrição.

🛠️ Configuração e Requisitos

1. Requisitos de Ambiente

Para rodar este script, você precisará ter o Python instalado e as seguintes bibliotecas:

pip install pandas reportlab openpyxl


2. Estrutura de Pastas

O projeto espera a seguinte estrutura de arquivos para funcionar corretamente:

/diretorio_do_projeto
├── catalogo_generator_alfabetico.py (O script principal)
├── produtos.xlsx (Planilha de dados)
├── logo_amaisciclo.png (Logo para Capa e Cabeçalho)
├── 10porcem.jpg (Imagem de desconto opcional)
└── img_produtos/
    ├── CODIGO.jpg (Imagens dos produtos, nomeadas pelo Código do Produto)
    └── CODIGO_LIMPO.png


3. Planilha de Dados (produtos.xlsx)

A planilha deve conter, no mínimo, as seguintes colunas para o script funcionar corretamente:

Coluna

Tipo de Dado

Descrição

Código do Produto

Texto

Código único usado para buscar a imagem (img_produtos/CODIGO.jpg). Obrigatório.

Descrição

Texto

Nome e detalhes do produto.

Preço Antigo

Numérico

Preço original do produto (opcional, será riscado).

Preço Promoção

Numérico

Preço em destaque (opcional, será exibido em vermelho).

Categoria

Texto

Categoria do produto (não usada para agrupamento na versão alfabética).

4. Configurações de Cores e Layout

No topo do arquivo catalogo_generator_alfabetico.py, você pode ajustar as cores e o layout:

# Cores
COR_AZUL_CODIGO = colors.Color(red=0.8, green=0.0, blue=0.0) # Vermelho Forte para o Código (e Preços Promocionais)
COR_FUNDO_ESCURO = colors.Color(red=0.0, green=0.0, blue=0.0) # Preto Absoluto para o fundo da capa
# ... outras cores


🚀 Como Executar o Projeto

Certifique-se de ter todos os requisitos instalados e a estrutura de arquivos correta (planilha, logo e pasta de imagens).

Abra o terminal ou prompt de comando na pasta do projeto.

Execute o script Python:

python catalogo_generator_alfabetico.py


Após a execução, o catálogo será gerado no caminho especificado por pdf_path.
