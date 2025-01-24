# Converto-Arquivos# Conversor de Arquivos para PDF, CSV e DOCX

Este é um script em Python que converte arquivos entre os formatos PDF, CSV, DOCX e TXT. Ele foi desenvolvido para facilitar a conversão entre esses formatos, salvando os resultados em uma pasta de saída.

## Funcionalidades

-   **Conversão para PDF:**
    -   Converte arquivos TXT, DOCX e CSV para PDF, preservando o texto e as quebras de linha.
-   **Conversão para CSV:**
    -   Converte arquivos TXT, PDF e DOCX para CSV, utilizando o texto bruto dos arquivos, com separador ";".
    -   Para outros formatos, utiliza o LibreOffice via linha de comando.
-   **Conversão para DOCX:**
    -   Converte arquivos TXT, PDF e CSV para DOCX. Os arquivos PDF são convertidos para DOCX, reconhecendo títulos, listas e parágrafos.

## Como Usar

1.  **Requisitos:**
    -   Python 3.6 ou superior instalado.
    -   Bibliotecas Python: `pymupdf`, `python-docx`, `reportlab`.
    -   LibreOffice instalado (necessário para conversão de formatos não-suportados para CSV).

2.  **Instalação das Bibliotecas:**
    Abra o terminal ou prompt de comando e execute:

    ```bash
    pip install pymupdf python-docx reportlab
    ```

3.  **Estrutura de Pastas:**
    -   Crie uma pasta chamada `Arquivo a modificar` (ou o nome que você definir em `PASTA_ENTRADA`).
    -   Coloque os arquivos que você deseja converter dentro desta pasta.
    -   O script criará uma pasta chamada `Arquivos modificados` (ou o nome que você definir em `PASTA_SAIDA`) para salvar os arquivos convertidos.

4.  **Execução do Script:**
    -   Salve o código Python como um arquivo (por exemplo, `convert_files.py`).
    -   Abra o terminal/prompt de comando, navegue até a pasta onde você salvou o script e execute:

    ```bash
    python convert_files.py
    ```

5.  **Arquivos de Saída:**
    -   Os arquivos convertidos (PDF, CSV e DOCX) estarão na pasta `Arquivos modificados`.
    -   Os nomes dos arquivos de saída são derivados dos nomes dos arquivos de entrada.

## Configurações

-   **`BASE_DIR`**: O diretório onde o script está localizado.
-   **`PASTA_ENTRADA`**: Pasta de entrada dos arquivos a serem convertidos (padrão: "Arquivo a modificar").
-   **`PASTA_SAIDA`**: Pasta onde os arquivos convertidos serão salvos (padrão: "Arquivos modificados").

## Funções Principais

-   **`Criacao_PDF_Texto(text, arquivo_saida)`**: Cria um arquivo PDF com o texto fornecido.
-   **`Leitura_Extracao(arquivo_entrada)`**: Lê um arquivo DOCX, CSV ou PDF e extrai o texto.
-   **`Convertor_PDF(arquivo_entrada, arquivo_saida)`**: Converte um arquivo TXT, DOCX ou CSV para PDF.
-   **`Convertor_CSV(arquivo_entrada, arquivo_saida)`**: Converte um arquivo para CSV.
-   **`Convertor_DOCX(arquivo_entrada, arquivo_saida)`**: Converte um arquivo para DOCX, detectando títulos e listas quando o formato de entrada é PDF.
-   **`main()`**: Função principal que processa todos os arquivos na pasta de entrada.

