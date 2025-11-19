# Curva A - Mercado Livre
Automatize a pesquisa de mercado no Mercado Livre. Um scraper com interface gráfica para extrair preços, dados de vendedores e avaliações a partir de uma lista em Excel.

Markdown

# 📈 Curva A - Scraper de Mercado Livre

Este é um scraper web com interface gráfica (GUI) que automatiza a coleta de dados de produtos no Mercado Livre. Ele foi projetado para auxiliar na análise de mercado, permitindo a extração de preços, informações de vendedores, avaliações e mais, a partir de uma lista de termos de busca.

## ✨ Funcionalidades

- **Coleta de Dados Automática**: Faz buscas no Mercado Livre a partir de uma lista de termos em um arquivo Excel.
- **Extração Detalhada**: Coleta dados tanto da página de busca quanto da página de detalhes do produto (PDP), incluindo:
  - Título, preço, tipo de anúncio (`Clássico` ou `Premium`), link.
  - Vendedor, quantidade de vendidos, nota média e número de avaliações.
- **Comparação de Preços**: Permite definir uma lista de "lojas próprias" para identificar se concorrentes estão vendendo produtos por um preço menor que o seu.
- **Interface Gráfica (GUI)**: Simplifica o uso da ferramenta com uma interface visual (Tkinter), sem a necessidade de usar a linha de comando.
- **Comportamento Humanizado**: Utiliza pausas aleatórias e simulação de rolagem de mouse para evitar detecção como um robô.
- **Saída em Excel**: Salva os dados coletados em um arquivo `.xlsx` de fácil visualização e análise.

## 🚀 Como Usar

### 1. Pré-requisitos

Certifique-se de que você tem o Python 3.7 ou superior instalado em sua máquina.

### 2. Instalação

Abra o seu terminal (ou Prompt de Comando/PowerShell) e execute o seguinte comando para instalar as bibliotecas necessárias:

```bash
pip install -r requirements.txt
3. Execução
Para iniciar a aplicação, simplesmente rode o script a partir do terminal:

python curva_a_ml.py
4. Usando a Interface
Arquivo Excel: Clique em Selecionar... para escolher o arquivo que contém os termos de busca na primeira coluna (Coluna A).

Opções: Ajuste as configurações de busca, como o número de resultados a capturar e se o navegador deve ser visível (Headless).

Lojas Específicas: Insira os nomes das lojas que você quer monitorar para fazer a comparação de preços. Separe os nomes por ponto e vírgula.

Pasta de Saída: Escolha onde os resultados serão salvos.

Iniciar: Clique em Iniciar para começar a coleta de dados. O log na parte inferior mostrará o progresso em tempo real.

📦 Como Empacotar (opcional)
Se você deseja criar um arquivo executável para o seu aplicativo (sem a necessidade de instalar Python ou as bibliotecas), você pode usar o PyInstaller.

Baixe o navegador Chromium do Playwright:

set PLAYWRIGHT_BROWSERS_PATH=ms-playwright
python -m playwright install chromium
Compile o executável:

pyinstaller --noconfirm --onedir --windowed ^
  --name "CurvaA-ML" ^
  --add-data "ms-playwright;ms-playwright" ^
  --hidden-import=playwright.sync_api --hidden-import=pyee ^
  curva_a_ml.py
O executável estará na pasta dist/CurvaA-ML/.

❤️ Apoie o Projeto

Este projeto foi desenvolvido com dedicação e tempo. Se esta ferramenta foi útil para você, considere fazer uma doação para me ajudar a continuar criando e
aprimorando projetos de código aberto.

Chave PIX 55df1ddb-4916-4cda-8a0e-fab0947764ca

https://buymeacoffee.com/douglas.onorio

Agradeço imensamente o seu apoio!
