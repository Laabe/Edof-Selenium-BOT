from edof.edof import Edof
from edof.manipulate_xl import reset_worksheet

try:
    reset_worksheet()
except:
    print('problem with the Excel file')

try:
    with Edof() as bot:
        bot.go_to_landing_page()
        bot.click_cookies_btn()
        bot.click_connexion_btn()
        bot.login()
        bot.go_to_all_folders_page()
        bot.itirate_throught_pages()
        bot.close()
except Exception as e:
    print(str(e))
