import os
from time import sleep
from edof.manipulate_xl import store_folder_in_excel
import edof.constants as const
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class Edof(webdriver.Chrome):
    def __init__(self, driver_path=r"/usr/bin/chromedriver"):
        self.driver_path = driver_path
        os.environ['PATH'] += self.driver_path
        options = webdriver.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--headless")
        options.add_argument("--remote-debugging-port=9222")
        options.add_experimental_option('excludeSwitches', ["enable-logging", "enable-automation"])
        super(Edof, self).__init__(options=options)
        self.delete_all_cookies()
        self.implicitly_wait(90)    
    
    def go_to_landing_page(self):
        self.get(const.LANDING_PAGE_URL)
        print('''============================= test session starts ==============================''')
        print('Entered website...')

    def click_cookies_btn(self):
        try:
            cookies_btn = WebDriverWait(self, 60).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "/html/body/div[2]/div[2]/div/mat-bottom-sheet-container/mcf-cm-banner/div[2]/button[1]"
                ))
            )
            # .click() to mimic button click
            cookies_btn.click()
            print('Clicked cookies btn')
        except Exception as e:
            print('Cookies btn problem')

    def click_connexion_btn(self):
        try:
            connexion_btn = WebDriverWait(self, 60).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    '//*[@id="landing-task-section-id"]/div[2]/a'
                ))
            )
            # .click() to mimic button 
            connexion_btn.click()
            print('Clicked connexion btn...')
        except Exception as e:
            print('connexion btn problem')

    def login(self):
        try:
            username = self.find_element(By.ID, 'username')
            username.send_keys(const.linkedin_username)
            print('username entered')
        except Exception as e:
            print('username input problem')
            
        # locate password form By_id

        try:
            password = self.find_element(By.ID, 'password')
            password.send_keys(const.linkedin_password)
            print('password entered entered')
        except Exception as e:
            print('password input problem')

        try:
            log_in_button = self.find_element(by=By.CSS_SELECTOR, value='button[name="submitBtn"]')
            # .click() to mimic button click
            log_in_button.click()
            print('Connecting...')
        except Exception as e:
            print('login btn problem')

    def go_to_all_folders_page(self):
        sleep(2)
        self.get("https://www.of.moncompteformation.gouv.fr/espace-prive/html/#/dossiers/tous")
        print('Connected')

    def itirate_through_folders_and_store(self):
        sleep(0.5)
        try:
            folders_table = WebDriverWait(self, 60).until(
                EC.visibility_of_element_located((
                    By.TAG_NAME, 
                    'tbody'
                ))
            )

            print('folders table found')
        except Exception as e:
            print('the Table is not found')
            print(str(e))
        
        try:
            folders_data = WebDriverWait(folders_table, 60).until(
                EC.visibility_of_all_elements_located((
                    By.TAG_NAME, 'tr'
                ))
            )
        except Exception as e:
            print('The table data is not found')
            print(str(e))
        
        for folder_data in folders_data:
            try:
                folder_number = WebDriverWait(folder_data, 60).until(
                    EC.presence_of_element_located((
                        By.CLASS_NAME, "color--primary-lighter"
                    ))
                )
                folder_number = folder_number.get_attribute("innerHTML")
                folder_number = folder_number.replace("n°", "")

                try:
                    folder_status = WebDriverWait(folder_data, 60).until(
                        EC.presence_of_element_located((
                            By.CLASS_NAME, "margin--smallest-bottom"
                        ))
                    )
                    folder_status = folder_status.get_attribute("innerHTML")
                    folder_status = folder_status.split("<")[0]
                    folder_status = folder_status.strip()

                except:
                    folder_status = 'Not found'
                    print('folder status not found')
                    break

            except:
                folder_number = 'Not found'
                print('folder number not found')
                break
            
            print('------------------------------------------------------------------------------')
            print(folder_number, folder_status)
            store_folder_in_excel(folder_number, folder_status, const.EXCEL_FILE)

    def itirate_throught_pages(self): 
        total_page = self.find_element(by=By.ID, value="pagination-total").get_attribute('innerHTML')

        for _ in range(1, int(total_page) + 1):
            self.itirate_through_folders_and_store()

            try:
                next_page = WebDriverWait(self, 60).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        '/html/body/sl7-app/sl7-registration-folders/main/div[1]/sl7-registration-folder-list/section/section/sl7-pagination/section/ul/li[7]'
                    ))
                )
                next_page.click()
                print('next Page', _ + 1)
            except Exception as e:
                print('next page problem')
                print(str(e))
                self.quit()

        print("""
            ===================================================================
                                            END
            ===================================================================
        """)     

    def close(self):
        self.quit()
