
# Import for the Web Bot
import selectors
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

from webdriver_manager.chrome import ChromeDriverManager

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


bot = WebBot()

def navigation():
    # Opens the BotCity website.
    bot.browse("https://www.lme.com/")

    bot.wait(5000)

    bot.find_element('//*[@id="meganav-drawer"]/nav/ul/li[4]/button', By.XPATH).click()

    bot.find_element('//*[@id="meganav-drawer"]/nav/ul/li[4]/div/ul/li[3]/button', By.XPATH).click()

    bot.find_element('//*[@id="meganav-drawer"]/nav/ul/li[4]/div/ul/li[3]/div/div[2]/ul/li[2]/a', By.XPATH).click()

    bot.find_element( '/html/body/header/div[3]/div/div/div/div[2]/div/a[4]', By.XPATH).click()

    bot.scroll_down(4)

    bot.find_element('/html/body/main/div/div[3]/div/div/div[2]/div[2]/div[2]/a', By.XPATH).click()

    bot.scroll_down(4)


def download_data():
    bot.find_element('/html/body/main/div/div[2]/div[1]/div/div/p[2]/a', By.XPATH).click()
    bot.wait(3000)
    bot.find_element('/html/body/main/div/div[2]/div[1]/div/div/p[3]/a', By.XPATH).click()
    bot.wait(3000)
    bot.find_element('/html/body/main/div/div[2]/div[1]/div/div/p[4]/a', By.XPATH).click()
    bot.wait(3000)





def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")


    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()


    
    # Development
    navigation()
    download_data()


 

    # Wait 3 seconds before closing
    bot.wait(3000)

   
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
