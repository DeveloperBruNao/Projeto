#Importando Bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import smtplib
import email.message
import openpyxl
# -------------------------------------------------------

# Configurarando a planilha Excel
excel_file = 'dados_site.xlsx'
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active
#----------------------------------------------


#Setando variaveis -------------------------------
XPATH_NomeCompleto = '//*[@id="169757284"]'
XPATH_Idade = '//*[@id="169757284"]'
XPATH_Especificacoes = '//*[@id="169757466"]'
XPATH_Estado =  '//*[@id="169757564"]'
XPATH_Emailcompleto = '//*[@id="169757828"]'
XPATH_Experiencia_Menos_Um = '//*[@id="169757784_1237474510_label"]/span[2]'
XPATH_Experiencia_Um = '//*[@id="169757784_1237474511_label"]/span[2]'
XPATH_Experiencia_Dois = '//*[@id="169757784_1237474512_label"]/span[2]'
XPATH_Superior_Cursando = '//*[@id="169758546_1237481630_label"]'
XPATH_Superior_Nao = '//*[@id="169758546_1237481628_label"]'
XPATH_Superior_Sim = '//*[@id="169758546_1237481629_label"]'
XPATH_Botao_Concluido  = '//*[@id="patas"]/main/article/section/form/div[2]/button'
#---------------------------------------------------

# Iterarando sobre as linhas da planilha (minrow=2 por causa que pula o cabeçalho)
for row in sheet.iter_rows(min_row=2, values_only=True):
    Nome_completo, Idade, Especificacoes_Python, Estado, Anos_experiência,Emai_completo, Ensino_superior = row

    # Inicializando o driver do Selenium
    driver = webdriver.Chrome()
    #------------------------------------

    # Abrindo a página do surveymonkey
    driver.get('https://pt.surveymonkey.com/r/58ZHB66')
    #---------------------------------------------------

    #Função que contêm parametros do xpath e valor do nomecompleto
    def preencher_NomeCompleto(xpath, valor):
        campo_selecionado = driver.find_element(By.XPATH, xpath)
        campo_selecionado.send_keys(valor)
    
    #Função que contêm parametros do xpath e valor da Idade
    def preencher_Idade(xpath,valor):
        campo_selecionado = driver.find_element(By.XPATH,xpath)
        campo_selecionado.send_keys(valor)
    
    #Função que contêm parametros do xpath e valor das especificações
    def preencher_Especificacoes(xpath,valor):
        campo_selecionado = driver.find_element(By.XPATH,xpath)
        campo_selecionado.send_keys(valor)

    #Função que contêm parametros do xpath e valor do Estado
    def preencher_Estado(xpath,valor):
        campo_selecionado = driver.find_element(By.XPATH,xpath)
        campo_selecionado.send_keys(valor)
    
    #Função que contêm parametros do xpath e valor do preencherEmail
    def preencher_Emailcompleto(xpath,valor):
        campo_selecionado = driver.find_element(By.XPATH,xpath)
        campo_selecionado.send_keys(valor)
        

    #Chamando as funções e passando os valores que estão dentro das variaveis
    preencher_NomeCompleto(XPATH_NomeCompleto, Nome_completo)
    preencher_Idade(XPATH_Idade, Idade)
    preencher_Especificacoes(XPATH_Especificacoes, Especificacoes_Python)
    preencher_Estado(XPATH_Estado, Estado)
    preencher_Emailcompleto(XPATH_Emailcompleto, Emai_completo)
    #------------------------------------------------------------------------

    #Função de clicar no botão desejado de acordo com a linha no Excel através de Condicionais
    def Anos_experiencia():
        if Anos_experiência == -1:
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Experiencia_Menos_Um)).click()
        elif Anos_experiência == 1:
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Experiencia_Um)).click()
        elif Anos_experiência == 2:
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Experiencia_Dois)).click()
        else:
            print('Ocorreu um erro')
    Anos_experiencia()
    #-----------------------------------------------------------------------------------------
    
    #Função de clicar no botão desejado de acordo com a linha no Excel através de Condicionais
    def Ensino_Superior():
        if Ensino_superior == 'Cursando':
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Superior_Cursando)).click()
        elif Ensino_superior == 'Nao':
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Superior_Nao)).click()
        elif Ensino_superior == 'Sim':
            campo_selecionado = driver.find_element(By.XPATH,(XPATH_Superior_Sim)).click()
        else:
            print('Deu erro')
    Ensino_Superior()
    #--------------------------------------------------------------------------------------------------

    #Função de botao para finalizar o formulário
    def Botao_Concluido():
        campo_selecionado = driver.find_element(By.XPATH,(XPATH_Botao_Concluido)).click()
    time.sleep(2)
    #-----------------------------------------------------------------------------------------

    #Variaveis para utilizar na configuração do email
    Subject = 'Python Projects'
    From = ''  #Colocar o seu email
    To = Emai_completo
    mail = '' #Colocar o seu email
    senha = '' #Colocar sua senha
    #-----------------------------------------------------

    # Configurar e enviar e-mail
    def Enviar_email():
        msg = email.message.Message()
        msg['Subject'] = Subject
        msg['From'] = From
        msg['To'] = To
        msg.set_payload(f'Olá {Nome_completo}, Seu formulário foi preenchido com sucesso!', charset='utf-8')
        
        #Variaveis para configuração do servidor SMTP
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        # Conectando ao servidor de e-mail e enviar e-mail
        smtp_server = smtp_server
        smtp_port = smtp_port
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(mail,senha)
        server.sendmail(msg['From'], [msg['To']], msg.as_bytes())
        server.quit()
        #----------------------------------------------------------------------------------------------
    # Registrar a ação no log
    with open('log.txt', 'a') as log_file:
        log_file.write(f'Ação realizada para {Nome_completo}: Preenchimento do formulário e envio de e-mail\n')
        driver.close()

    continue
# Fechar o navegador e salvar a planilha Excel
driver.quit()
time.sleep(2)
workbook.save(excel_file)

