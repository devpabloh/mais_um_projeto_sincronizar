from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import keyboard

class VcardSync:
    def __init__(self, driver_path, base_url, username, password):
        self.base_url = base_url
        self.username = username
        self.password = password
        self.driver = webdriver.Chrome(executable_path=driver_path)
        self.login()

    def login(self):
        self.driver.get(self.base_url)
        time.sleep(2)
        username_field = self.driver.find_element(By.ID, "pablo.henrique")
        password_field = self.driver.find_element(By.ID, "@Taisatt84671514")
        username_field.send_keys(self.username)
        password_field.send_keys(self.password)
        password_field.send_keys(Keys.RETURN)
        time.sleep(3)

    def import_vcard(self, file_path):
        self.driver.find_element(By.ID, "calendarid").click()
        time.sleep(3)
        #localizando o campo de upload
        file_input = self.driver.find_element(By.ID, '//input[@value="importar"]')
        file_input.send_keys(file_path)
        time.sleep(3)
        submit_button = self.driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        time.sleep(3)
        print("Arquivo vCard importado com sucesso.")

    def create_event(self, event_data):
        #Criando evento via formulário web
        self.driver.find_element(By.ID, "calendarid").click()
        time.sleep(3)
        summary_field = self.driver.find_element(By.XPATH, "//img[@title='Novo evento']").click()
        time.sleep(3)
        start_field = self.driver.find_element(By.ID, "start[str]")
        end_field = self.driver.find_element(By.ID, "end[str]")
        location_field = self.driver.find_element(By.XPATH, 'input[@name="cal[location]"]')
        description_field = self.driver.find_element(By.XPATH, 'textarea[@name="cal[description]"]')

        # preenchendo os campos do formulário com os dados do evento
        summary_field.send_keys(event_data.get('summary', 'Sem título'))
        start_field.send_keys(event_data.get('start', ''))
        end_field.send_keys(event_data.get('end', ''))
        location_field.send_keys(event_data.get('location', ''))
        description_field.send_keys(event_data.get('description', ''))

        submit_button = self.driver.find_element(By.ID, "submit_button").click()
        time.sleep(3)
        print("Evento criado com sucesso.")

    def update_event(self, event_id, event_data):
        #Atualizando evento via formulário web
        self.driver.get(f"{self.base_url}/index.php?menuaction=calendar.uicalendar.view&cal_id={event_id}&date={date}")
        time.sleep(3)
        # localizando o botão de edição
        button_edit = self.driver.find_element(By.XPATH, '//input[@value="Editar"]').click()
        # atualizando os campos 
        summary_field = self.driver.find_element(By.XPATH, "//img[@title='Novo evento']").click()
        start_field = self.driver.find_element(By.ID, "start[str]")
        end_field = self.driver.find_element(By.ID, "end[str]")
        location_field = self.driver.find_element(By.XPATH, 'input[@name="cal[location]"]')
        time.sleep(3)

        summary_field.send_keys(event_data.get('summary', 'Sem título'))
        start_field.send_keys(event_data.get('start', ''))
        end_field.send_keys(event_data.get('end', ''))
        location_field.send_keys(event_data.get('location', ''))
        description_field.send_keys(event_data.get('description', ''))

        submit_button = self.driver.find_element(By.ID, "submit_button").click()
        time.sleep(3)

        print(f"Evento {event_id} atualizado com sucesso.")

    def delete_event(self, event_id):
        #Deletando evento via formulário web
        self.driver.get(f"{self.base_url}/index.php?menuaction=calendar.uicalendar.view&cal_id={event_id}&date={date}")
        time.sleep(3)
        delete_button = self.driver.find_element(By.XPATH, '//input[@value="Remover"]').click()
        time.sleep(3)
        keyboard.press_and_release("enter")
        time.sleep(3)
        print(f"Evento {event_id} deletado com sucesso.")
        

    def close(self):
        self.driver.quit()

        
        
