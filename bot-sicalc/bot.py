"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = DesktopBot()
    #bot.browse("http://www.botcity.dev")

    bot.execute(r"C:\Program Files (x86)\Programas RFB\Sicalc Auto Atendimento\SicalcAA.exe")
    
    if not bot.find( "popup-esclarecimento", matching=0.97, waiting_time=10000):
        not_found("popup-esclarecimento")
    bot.click_relative(218, 261)
    
    if not bot.find( "funcoes", matching=0.97, waiting_time=10000):
        not_found("funcoes")
    bot.click()
    
    if not bot.find( "preenchimento-darf", matching=0.97, waiting_time=10000):
        not_found("preenchimento-darf")
    bot.click()
    
    if not bot.find( "codigo-receita", matching=0.97, waiting_time=10000):
        not_found("codigo-receita")
    bot.click_relative(196, 12)

    bot.kb_type("5629")

    bot.tab()

    bot.wait(2000)
    
    if not bot.find( "pa", matching=0.97, waiting_time=10000):
        not_found("pa")
    bot.click_relative(40, 33)
    
    bot.paste("310120")
    
    if not bot.find( "valor-reais", matching=0.97, waiting_time=10000):
        not_found("valor-reais")
    bot.click_relative(51, 36)

    bot.paste("1000")

    if not bot.find( "calcular", matching=0.97, waiting_time=10000):
        not_found("calcular")
    bot.click()
    
    if not bot.find( "botao-darf", matching=0.97, waiting_time=10000):
       not_found("botao-darf")
    bot.click()
   
    
    if not bot.find( "nome", matching=0.97, waiting_time=10000):
        not_found("nome")
    bot.click_relative(94, 31)

    bot.paste("Petrobras")
    
    if not bot.find( "telefone", matching=0.97, waiting_time=10000):
        not_found("telefone")
    bot.click_relative(64, 31)
    
    bot.paste("1199991234")
    
    if not bot.find( "cnpj", matching=0.97, waiting_time=10000):
        not_found("cnpj")
    bot.click_relative(200, 11)

    bot.paste("33000167000101")

    if not bot.find( "referencia", matching=0.97, waiting_time=10000):
        not_found("referencia")
    bot.click_relative(197, 12)
    
    bot.paste("0")
    
    if not bot.find( "imprimir", matching=0.97, waiting_time=10000):
        not_found("imprimir")
    bot.click()
    
    if not bot.find( "saida-imprimir", matching=0.97, waiting_time=10000):
        not_found("saida-imprimir")
    bot.click_relative(198, 463)
    

    bot.paste(r"C:\Users\sabri\Desktop\TreinamentoBotcity\bot-sicalc\pdfs\Darf1.pdf")
    bot.enter()

    bot.wait(2000)

    bot.alt_f4()

    bot.alt_f4()
    
    

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()


