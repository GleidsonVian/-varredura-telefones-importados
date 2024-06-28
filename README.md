# Varredor de Celulares

Este projeto realiza a varredura de um site de telefones importados, coleta informações dos produtos listados e exporta esses dados para um arquivo Excel. Após a exportação, o arquivo é enviado por email automaticamente.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `selenium`
  - `pandas`
  - `pywin32`
- Navegador Chrome e ChromeDriver compatível com a versão do Chrome instalada

## Instalação

1. Clone este repositório:
    ```sh
    git clone https://github.com/seu-usuario/varredor-de-celulares.git
    cd varredor-de-celulares
    ```

2. Instale as dependências:
    ```sh
    pip install -r requirements.txt

    ```

3. Baixe e configure o ChromeDriver:
    - Faça o download do ChromeDriver em: [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)
    - Extraia o arquivo baixado e mova-o para um diretório que esteja no PATH do sistema, ou especifique o caminho completo no código Python.

## Uso

1. Edite o caminho do arquivo Excel e o endereço de email no código, conforme necessário:
    ```python
    # Caminho do arquivo Excel
    self.df.to_excel(r'C:\caminho\para\celulares.xlsx', index=False)
    
    # Endereço de email
    mail.To = 'seu-email@exemplo.com'
    ```

2. Execute o script:
    ```sh
    python bot.py
    ```

## Estrutura do Projeto

- `bot.py`: Script principal que contém a classe `Varredor_de_celulares` e o fluxo do programa.
- `celulares.xlsx`: Arquivo gerado contendo a lista de celulares.

## Funcionamento

1. **Instanciar site**: O navegador Chrome é aberto e o site alvo é acessado.
2. **Varrer site**: O script navega pelas páginas do site, coletando os nomes dos produtos.
3. **Exportar para Excel**: Os dados coletados são exportados para um arquivo Excel.
4. **Enviar email**: O arquivo Excel é enviado por email para o destinatário especificado.

## Contribuições

Contribuições são bem-vindas! Por favor, envie um pull request ou abra uma issue para discutir as mudanças que você gostaria de fazer.

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.
