from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import win32com.client as win32
import time

class Varredor_de_celulares():
    # Classe para varrer um site de celulares, exportar os dados e enviar por email

    def iniciar_programa(self):
        # Método principal que chama os outros métodos em sequência
        self.instanciar_site()
        self.varrer_site()
        self.exportar_para_excel()
        self.enviar_email()

    def instanciar_site(self):
        # Inicia o navegador e abre o site
        self.driver = webdriver.Chrome()
        self.driver.get("https://telefonesimportados.netlify.app/")
        self.lista_de_produtos = []
        self.lista_de_precos = []
        
    def varrer_site(self):
        # Coleta nomes e preços dos produtos em todas as páginas do site
        while True:
            try:
                # Localiza os nomes dos produtos e preços na página
                nome_produtos = self.driver.find_elements(By.XPATH, "//div[@class='single-shop-product']//h2")
                precos = self.driver.find_elements(By.XPATH, "//div[@class='product-carousel-price']//ins")
                
                # Armazena os textos dos produtos e preços nas listas
                for produto in nome_produtos:
                    self.lista_de_produtos.append(produto.text)
                for preco in precos:
                    self.lista_de_precos.append(preco.text)

                # Tenta encontrar o botão de próxima página
                botao_proximo = self.driver.find_element(By.XPATH, "//a[@aria-label='Next']")
                
                if not botao_proximo.is_enabled():
                    # Sai do loop se o botão "Próximo" não estiver habilitado
                    break
                botao_proximo.click()  # Clica no botão para ir à próxima página
                time.sleep(2)  # Espera 2 segundos para a próxima página carregar
                
                print('Navegando para a próxima página')
                
            except Exception as e:
                # Captura e imprime qualquer erro
                print(f"Erro: {e}")
                break

    def exportar_para_excel(self):
        # Exporta os dados coletados para um arquivo Excel
        if len(self.lista_de_produtos) != len(self.lista_de_precos):
            print("Erro: Listas de produtos e preços têm tamanhos diferentes")
            return

        # Cria um DataFrame com os produtos e preços
        self.df = pd.DataFrame({
            'Celulares': self.lista_de_produtos,
            'Preços': self.lista_de_precos
        })
        
        # Exporta o DataFrame para um arquivo Excel
        self.df.to_excel(r'C:\Users\gleid\pasta4\Python\varredura telefones importados\celulares.xlsx', index=False)
        print('Dados exportados para Excel')

    def enviar_email(self):
        # Envia um email com o arquivo Excel em anexo
        try:
            # Cria uma instância do Outlook
            outlook = win32.Dispatch('outlook.application')

            # Cria um novo item de email
            mail = outlook.CreateItem(0)

            # Define o destinatário do email
            mail.To = 'gleidson.testes1@outlook.com'

            # Define o assunto do email
            mail.Subject = 'Tabela de Celulares'

            # Define o corpo do email com a tabela em HTML
            mail.HTMLBody = f"""
            <p>Segue a lista de celulares varridos</p>
            {self.df.to_html()}
            """

            # Anexa o arquivo Excel
            anexo = r"C:\Users\gleid\pasta4\Python\varredura telefones importados\celulares.xlsx"
            mail.Attachments.Add(anexo)
            
            # Envia o email
            mail.Send()
            print('Email enviado!')
        except Exception as e:
            # Captura e imprime qualquer erro ao enviar o email
            print(f"Erro ao enviar email: {e}")

# Cria uma instância da classe e inicia o programa
programa = Varredor_de_celulares()
programa.iniciar_programa()
