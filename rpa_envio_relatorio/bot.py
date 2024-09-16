"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/web/
"""


# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from webdriver_manager.chrome import ChromeDriverManager
from botcity.plugins.http import BotHttpPlugin
from botcity.plugins.email import BotEmailPlugin
import PyPDF2
import re
import pandas as pd

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def extrair_pdf():

    path_arquivo = r"docs\Controle_SUS.pdf"

    with open(path_arquivo, "rb") as arquivo:
        leitor_pdf = PyPDF2.PdfReader(arquivo)

        texto = " "

        for pagina in leitor_pdf.pages:
            texto += pagina.extract_text()

        texto = re.sub(r'(\w+\s+\w+\s+\d{3}\.\d{3}\.\d{3}-\d{2}\s+(Sim|Não))', r'\1\n', texto)

        linhas = texto.splitlines()

        linhas_com_nao = []

        for linha in linhas:
            if "Não" in linha:
                linhas_com_nao.append(linha.strip())


        coluna1 = []
        coluna2 = []
        
        for linha in linhas_com_nao:

            corresponder = re.match(r'(\w+)\s+(\w+)\s+(\d{3}\.\d{3}\.\d{3}-\d{2})\s+(Sim|Não)', linha)
            if corresponder:
                nome_completo = f"{corresponder.group(1)} {corresponder.group(2)}"
                coluna1.append(nome_completo)  
                coluna2.append(corresponder.group(3))  
   

        return pd.DataFrame({
            'Nome': coluna1,
            'CPF': coluna2,
        })



def salvar_dados_excel():

    dados = extrair_pdf()

    df = pd.DataFrame(dados)

    nome_arquivo = r"docs\RelatorioSUS.xlsx"

    df.to_excel(nome_arquivo, index=False)

    print(f"Dados Salvos em {nome_arquivo}")


def enviar_email(user_email, user_senha, to_email, assunto, conteudo, arquivo_path ):
    
    email = BotEmailPlugin()

    email.configure_imap("imap.gmail.com", 993)

    email.configure_smtp("smtp.gmail.com", 587)

    email.login(user_email, user_senha)

    email.send_message(assunto, conteudo, to_email, attachments=arquivo_path, use_html=True)
     
    email.disconnect()

    print("E-mail enviado com sucesso!")



def parametro_emails():

    to = ["sabrina.frazao@ifam.edu.br", "sabrinadasilvafrazao@gmail.com"]
    subject = "Relatorio_SUS"
    body = "Relatorio atualizado"
    files = ["docs\RelatorioSUS.xlsx"]


    enviar_email(
        user_email="sabrinadasilvafrazao@gmail.com",
        user_senha="wsoc rayq oueq weiy",
        to_email=to,
        assunto=subject,
        conteudo=body,
        arquivo_path=files
    )


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.browse("https://www.botcity.dev")



    salvar_dados_excel()

    parametro_emails()


    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
