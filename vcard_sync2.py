from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse, parse_qs
import time
import keyboard

# Função para formatar a data
def formatar_data(data_str):
    # More robust date formatting
    if not data_str:
        return ''
    
    # Handle YYYYMMDD format
    if len(data_str) == 8:
        ano = data_str[:4]
        mes = data_str[4:6]
        dia = data_str[6:8]
        return f"{dia}/{mes}/{ano}"
    
    # Try to handle other formats or return original
    return data_str

class sincronizarExpresso:
    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.driver = None

    def login(self):
        self.driver = webdriver.Chrome()
        self.driver.get("https://www.expresso.pe.gov.br/login.php?cd=1")
        time.sleep(3)
        inputlogin = self.driver.find_element(By.XPATH, "//input[@name='user']")
        inputlogin.clear()
        inputlogin.send_keys(self.username)
        inputSenha = self.driver.find_element(By.XPATH, "//input[@type='password']")
        inputSenha.clear()
        inputSenha.send_keys(self.password)
        botaoConectar = self.driver.find_element(By.XPATH, "//div[@class='botao-conectar']//input[@type='submit']")
        botaoConectar.click()
        time.sleep(3)
    
    def selecionarCalendario(self):
        botaoCalendario = self.driver.find_element(By.XPATH, "//a[@href='/calendar/index.php']//img[@id='calendarid']")
        botaoCalendario.click()
        time.sleep(3)
        calendarioEsteMes = self.driver.find_element(By.XPATH, "//img[@title='Este mês']")
        calendarioEsteMes.click()
        time.sleep(3)
        
    def obterEventos(self):
        # Adicionar debug para ver a URL atual
        print(f"URL atual: {self.driver.current_url}")
        
        # Aumentar tempo de espera para garantir que a página carregue completamente
        time.sleep(5)
        
        try:
            # Esperar até que os elementos de eventos estejam presentes
            wait = WebDriverWait(self.driver, 15)
            
            # Tentar encontrar os eventos com diferentes seletores
            try:
                eventos = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "event_entry")))
                print(f"Encontrados {len(eventos)} eventos com class='event_entry'")
            except:
                print("Não foi possível encontrar elementos com class='event_entry'")
                # Tentar encontrar elementos div que contêm eventos
                eventos = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@id, 'calendar_event_entry')]")))
                print(f"Encontrados {len(eventos)} eventos com div[id*='calendar_event_entry']")
            
            # Salvar HTML da página para análise
            with open("pagina_calendario.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
            print("HTML da página salvo em pagina_calendario.html para análise")
            
            eventos_lista = []
            
            # Iterar sobre os eventos
            for i, evento in enumerate(eventos):
                print(f"\nProcessando evento {i+1}:")
                
                # Extrair URL completa
                try:
                    # Tentar diferentes abordagens para encontrar o link
                    try:
                        tag_url = evento.find_element(By.TAG_NAME, "a")
                    except:
                        # Se o próprio evento é uma tag <a>
                        tag_url = evento if evento.tag_name == "a" else None
                    
                    if tag_url:
                        url_completa = tag_url.get_attribute("href")
                        print(f"URL encontrada: {url_completa}")
                    else:
                        url_completa = ""
                        print("Não foi possível encontrar URL")
                except Exception as e:
                    url_completa = ""
                    tag_url = None
                    print(f"Erro ao extrair URL: {e}")
                
                # Extrair horário
                try:
                    span_horario = evento.find_element(By.CSS_SELECTOR, "font span")
                    horario_texto = span_horario.text.strip()
                    horarios = horario_texto.split('-')
                    horario_inicio = horarios[0].strip() if len(horarios) > 0 else ""
                    horario_fim = horarios[1].strip() if len(horarios) > 1 else ""
                except Exception as e:
                    horario_inicio = ""
                    horario_fim = ""
                    print(f"Erro ao extrair horário: {e}")
                
                # Extrair título
                try:
                    titulo_el = evento.find_element(By.TAG_NAME, "b")
                    titulo = titulo_el.text.strip()
                except Exception as e:
                    titulo = ""
                    print(f"Erro ao extrair título: {e}")
                
                # Extrair descrição
                try:
                    descricao_el = evento.find_element(By.TAG_NAME, "i")
                    descricao = descricao_el.text.strip()
                except Exception as e:
                    descricao = ""
                    print(f"Erro ao extrair descrição: {e}")
                
                # Extrair participantes
                try:
                    imagens = evento.find_elements(By.TAG_NAME, "img")
                    participantes = imagens[1].get_attribute("title") if len(imagens) >= 2 else ""
                except Exception as e:
                    participantes = ""
                    print(f"Erro ao extrair participantes: {e}")

                # Parsear URL para obter id e data
                id_evento = ""
                data = ""
                data_formatada = ""
                if url_completa:
                    try:
                        parsed_url = urlparse(url_completa)
                        params = parse_qs(parsed_url.query)
                        print(f"Parâmetros da URL: {params}")
                        id_evento = params.get("cal_id", [""])[0] if "cal_id" in params else ""
                        data = params.get("date", [""])[0] if "date" in params else ""
                        data_formatada = formatar_data(data)
                        print(f"ID extraído: {id_evento}")
                        print(f"Data extraída: {data} -> {data_formatada}")
                    except (IndexError, KeyError) as e:
                        print(f"Erro ao processar URL: {e}")

                # Criar dicionário do evento
                evento_info = {
                    "id": id_evento,
                    "data": data_formatada,
                    "inicio": horario_inicio,
                    "fim": horario_fim,
                    "titulo": titulo,
                    "descricao": descricao,
                    "url": url_completa,
                    "participantes": participantes,
                    "tag_url": tag_url
                }
                
                eventos_lista.append(evento_info)
            
            return eventos_lista
            
        except Exception as e:
            print(f"Erro geral ao obter eventos: {e}")
            return []
    
    def fechar(self):
        if self.driver:
            self.driver.quit()

# Exemplo de uso
if __name__ == "__main__":
    sync = sincronizarExpresso("pablo.henrique1", "@Taisatt84671514")
    sync.login()
    sync.selecionarCalendario()
    eventos = sync.obterEventos()
    
    # Exibir informações dos eventos
    for evento in eventos:
        # Evitar erro ao imprimir tag_url que é um objeto WebElement
        tag_url_str = "WebElement encontrado" if evento["tag_url"] else "Não encontrado"
        
        print("Tag da URL:", tag_url_str)
        print("Id:", evento["id"])
        print("Data:", evento["data"])
        print("Início:", evento["inicio"])
        print("Fim:", evento["fim"])
        print("Título:", evento["titulo"])
        print("Descrição:", evento["descricao"])
        print("URL completa:", evento["url"])
        print("Participantes:", evento["participantes"])
        print("---")
    
    sync.fechar()
