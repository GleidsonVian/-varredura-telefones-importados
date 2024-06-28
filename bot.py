from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import win32com.client as win32

class Varredor_de_celulares():

    def iniciar_programa(self):
        self.instanciar_site()
        self.varrer_site()
        self.exportar_para_excel()
        self.enviar_email()

    def instanciar_site(self):
        self.driver = webdriver.Chrome()
        self.driver.get("https://telefonesimportados.netlify.app/")
        self.lista_de_produtos = []
        
    
    def varrer_site(self):
        while True:
            try:
                nome_produtos = self.driver.find_elements(By.XPATH, "//div[@class='single-shop-product']//h2")
                precos = self.driver.find_elements(By.XPATH, "//div[@class='product-carousel-price']//ins")
               
                for produto in nome_produtos:
                    self.lista_de_produtos.append(produto.text)
                for preco in precos:
                    self.lista_de_precos.append(preco.text)

                botao_proximo = self.driver.find_element(By.XPATH, "//a[@aria-label='Next']")
                
                if not botao_proximo.is_enabled():
                    break
                botao_proximo.click()
                print('Navegando para proxima pagina')
                
            except Exception:
                print('')
                break

    def exportar_para_excel(self):
        if len(self.lista_de_produtos) != len(self.lista_de_precos):
            print("Erro: Listas de produtos e preços têm tamanhos diferentes")
            return

        self.df = pd.DataFrame({
            'Celulares': self.lista_de_produtos,
            'Preços': self.lista_de_precos
        })
        
        self.df.to_excel(r'celulares.xlsx', index=False)
        print('Dados exportados para Excel')

    def enviar_email(self):
        # Cria uma instância do Outlook
        outlook = win32.Dispatch('outlook.application')

        # Cria um novo item de email
        mail = outlook.CreateItem(0)

        # Define o destinatário do email
        mail.To = 'seuemail@outlook.com'

        # Define o assunto do email
        mail.Subject = 'Tabela de Celulares'

        # Define o corpo do email
        mail.HTMLBody = f"""

        <p>Segue a lista de celulares varridos</p>
        {self.df.to_html()}
        """
        anexo = r"C:caminho do arquivo na sua maquina\celulares.xlsx"
        mail.Attachments.Add(anexo)
        # Envia o email
        mail.Send()
        

        # Imprime mensagem de confirmação no console
        print('Email enviado!')

programa = Varredor_de_celulares()
programa.iniciar_programa()


