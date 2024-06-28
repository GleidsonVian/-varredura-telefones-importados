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
                for produto in nome_produtos:
                    self.lista_de_produtos.append(produto.text)

                botao_proximo = self.driver.find_element(By.XPATH, "//a[@aria-label='Next']")
                
                if not botao_proximo.is_enabled():
                    break
                botao_proximo.click()
                print('Navegando para proxima pagina')
                
            except Exception:
                print('')
                break

    def exportar_para_excel(self):
        self.df = pd.DataFrame(self.lista_de_produtos, columns=['Celulares'])
        self.df.to_excel(r'C:\Users\gleid\pasta4\Python\varredura telefones importados\celulares.xlsx', index=False)

    def enviar_email(self):
        # Cria uma instância do Outlook
        outlook = win32.Dispatch('outlook.application')

        # Cria um novo item de email
        mail = outlook.CreateItem(0)

        # Define o destinatário do email
        mail.To = 'gleidson.testes1@outlook.com'

        # Define o assunto do email
        mail.Subject = 'Tabela de Celulares'

        # Define o corpo do email
        mail.HTMLBody = f"""

        <p>Segue a lista de celulares varridos</p>
        {self.df.to_html()}
        """
        anexo = r"C:\Users\gleid\pasta4\Python\varredura telefones importados\celulares.xlsx"
        mail.Attachments.Add(anexo)
        # Envia o email
        mail.Send()
        

        # Imprime mensagem de confirmação no console
        print('Email enviado!')

programa = Varredor_de_celulares()
programa.iniciar_programa()


