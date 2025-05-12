from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urlparse, parse_qs
import time
import keyboard
from datetime import datetime, timedelta
import re



# Função para formatar a data
def formatar_data(data_str):
    # More robust date formatting
    if not data_str:
        return ""

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
        try:
            # Configurações do Chrome
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            
            # Instalar o ChromeDriver usando WebDriver Manager
            driver_path = ChromeDriverManager().install()
            service = Service(driver_path)
            
            # Inicializar o driver com as opções
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # Configurar timeout
            self.driver.set_page_load_timeout(60)
            
            # Acessar a página
            self.driver.get("https://www.expresso.pe.gov.br/login.php?cd=1")
            time.sleep(5)  # Esperar um pouco mais para a página carregar
            
            # Resto do código de login...
            inputlogin = self.driver.find_element(By.XPATH, "//input[@name='user']")
            inputlogin.clear()
            inputlogin.send_keys(self.username)
            inputSenha = self.driver.find_element(By.XPATH, "//input[@type='password']")
            inputSenha.clear()
            inputSenha.send_keys(self.password)
            botaoConectar = self.driver.find_element(
                By.XPATH, "//div[@class='botao-conectar']//input[@type='submit']"
            )
            botaoConectar.click()
            time.sleep(5)
            
        except Exception as e:
            print(f"Erro ao inicializar o Chrome: {e}")
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()
            raise

    def selecionarCalendario(self):
        botaoCalendario = self.driver.find_element(
            By.XPATH, "//a[@href='/calendar/index.php']//img[@id='calendarid']"
        )
        botaoCalendario.click()
        time.sleep(3)
        calendarioEsteMes = self.driver.find_element(
            By.XPATH, "//img[@title='Este mês']"
        )
        calendarioEsteMes.click()
        time.sleep(3)

    def obterEventos(self):
        time.sleep(5)

        try:
            # Esperar até que os elementos de eventos estejam presentes
            wait = WebDriverWait(self.driver, 30)

            # Tentar encontrar os eventos com diferentes seletores
            try:
                eventos = wait.until(
                    EC.presence_of_all_elements_located(
                        (
                            By.XPATH,
                            "//div[@id='calendar_event_entry']/a[@class='event_entry']",
                        )
                    )
                )
                print(f"Encontrados {len(eventos)} eventos com class='event_entry'")
            except:
                print("Não foi possível encontrar elementos com class='event_entry'")
                # Tentar encontrar elementos div que contêm eventos
                eventos = wait.until(
                    EC.presence_of_all_elements_located(
                        (By.XPATH, "//div[contains(@id, 'calendar_event_entry')]")
                    )
                )
                print(
                    f"Encontrados {len(eventos)} eventos com div[id*='calendar_event_entry']"
                )

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

                # Extrair horário - versão mais robusta
                try:
                    # Tentar várias abordagens para encontrar o horário
                    # Abordagem 1: Buscar diretamente por span com cor preta
                    try:
                        span_horario = evento.find_element(
                            By.CSS_SELECTOR, "span[style*='color: black']"
                        )
                        horario_texto = span_horario.text.strip()
                    except:
                        # Abordagem 2: Buscar font e depois span
                        try:
                            font_tag = evento.find_element(By.TAG_NAME, "font")
                            span_horario = font_tag.find_element(By.TAG_NAME, "span")
                            horario_texto = span_horario.text.strip()
                        except:
                            # Abordagem 3: Tentar extrair de todo o texto da tag font
                            try:
                                font_tag = evento.find_element(By.TAG_NAME, "font")
                                texto_completo = font_tag.text
                                # Procurar por padrão de hora (HH:MM-HH:MM)
                                padrao_hora = re.search(
                                    r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2})", texto_completo
                                )
                                if padrao_hora:
                                    horario_inicio = padrao_hora.group(1)
                                    horario_fim = padrao_hora.group(2)
                                else:
                                    horario_inicio = ""
                                    horario_fim = ""
                            except:
                                horario_inicio = ""
                                horario_fim = ""

                except Exception as e:
                    # Bloco except para o try principal
                    horario_inicio = ""
                    horario_fim = ""
                    print(f"Erro ao extrair horário: {e}")

                # Se encontrou o texto de horário no formato "14:00-16:00"
                if "horario_texto" in locals() and "-" in horario_texto:
                    horarios = horario_texto.split("-")
                    horario_inicio = horarios[0].strip()
                    horario_fim = horarios[1].strip()

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
                    participantes = (
                        imagens[1].get_attribute("title") if len(imagens) >= 2 else ""
                    )
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
                        id_evento = (
                            params.get("cal_id", [""])[0] if "cal_id" in params else ""
                        )
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
                    "tag_url": tag_url,
                }

                print(f"Horário extraído: início={horario_inicio}, fim={horario_fim}")

                eventos_lista.append(evento_info)

            return eventos_lista

        except Exception as e:
            print(f"Erro geral ao obter eventos: {e}")
            return []

    def fechar(self):
        if self.driver:
            self.driver.quit()

    def create_event(self, event_data):
        # Verificar se o evento já existe antes de criar
        try:
            # Primeiro obter eventos existentes
            eventos = self.obterEventos()

            # Verificar se já existe um evento com o mesmo título e data/hora
            for evento in eventos:
                if evento.get("titulo", "") == event_data.get(
                    "titulo", ""
                ) and evento.get("data", "") == event_data.get("data", ""):
                    # Verificar horário com tolerância de 5 minutos
                    hora_evento = evento.get("inicio", "").split(":")

                    # Obter o horário de início do evento a ser criado
                    hora_novo = None
                    if "inicio" in event_data and event_data["inicio"]:
                        if (
                            isinstance(event_data["inicio"], str)
                            and ":" in event_data["inicio"]
                        ):
                            hora_novo = event_data["inicio"].split(":")
                        elif isinstance(event_data["inicio"], datetime):
                            hora_novo = [
                                str(event_data["inicio"].hour),
                                str(event_data["inicio"].minute),
                            ]
                    elif "hora_inicio" in event_data and event_data["hora_inicio"]:
                        if (
                            isinstance(event_data["hora_inicio"], str)
                            and ":" in event_data["hora_inicio"]
                        ):
                            hora_novo = event_data["hora_inicio"].split(":")
                        elif isinstance(event_data["hora_inicio"], datetime):
                            hora_novo = [
                                str(event_data["hora_inicio"].hour),
                                str(event_data["hora_inicio"].minute),
                            ]

                    if len(hora_evento) == 2 and hora_novo and len(hora_novo) == 2:
                        minutos_evento = int(hora_evento[0]) * 60 + int(hora_evento[1])
                        minutos_novo = int(hora_novo[0]) * 60 + int(hora_novo[1])

                        if abs(minutos_evento - minutos_novo) <= 5:
                            print(
                                f"Evento já existe no Expresso: {evento['titulo']} em {evento['data']} {evento['inicio']}"
                            )
                            # Retornar o evento existente em vez de criar um novo
                            return evento

            # Se não encontrou evento duplicado, continuar com a criação
            # Navegando até a página de criação de eventos
            self.driver.get(
                "https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.add&date=20250423"
            )
            time.sleep(10)

            # Preenchendo o formulário
            input_titulo = self.driver.find_element(
                By.XPATH, "//input[@name='cal[title]']"
            )
            input_titulo.clear()
            input_titulo.send_keys(event_data["titulo"])

            # Selecionando a descrição
            input_descricao = self.driver.find_element(
                By.XPATH, "//textarea[@name='cal[description]']"
            )
            input_descricao.clear()
            input_descricao.send_keys(event_data["descricao"])

            # Selecionando a localização
            input_de_licalizacao = self.driver.find_element(
                By.XPATH, "//input[@name='cal[location]']"
            )
            input_de_licalizacao.clear()
            input_de_licalizacao.send_keys(event_data.get("localizacao", ""))

            # Selecionando a data de inicio
            input_data = self.driver.find_element(
                By.XPATH, "//input[@name='start[str]']"
            )
            input_data.clear()
            input_data.send_keys(event_data["data"])

            # Selecionando a data de fim
            input_data_fim = self.driver.find_element(
                By.XPATH, "//input[@name='end[str]']"
            )
            input_data_fim.clear()
            input_data_fim.send_keys(event_data["data"])

            # Inicializar as variáveis de horário com valores padrão
            hora_inicio = "00"
            minuto_inicio = "00"
            hora_fim = "00"
            minuto_fim = "00"

            # Verificar se é um evento de dia inteiro
            if event_data.get("dia_inteiro", False):
                # Para eventos de dia inteiro
                hora_inicio = "00"
                minuto_inicio = "00"
                hora_fim = "23"
                minuto_fim = "59"
            else:
                # Para eventos com horário específico - verificar múltiplos campos possíveis
                # Primeiro tentar 'inicio'
                if "inicio" in event_data and event_data["inicio"]:
                    if (
                        isinstance(event_data["inicio"], str)
                        and ":" in event_data["inicio"]
                    ):
                        hora_inicio, minuto_inicio = event_data["inicio"].split(":")
                    elif isinstance(event_data["inicio"], datetime):
                        hora_inicio = str(event_data["inicio"].hour).zfill(2)
                        minuto_inicio = str(event_data["inicio"].minute).zfill(2)
                # Depois tentar 'hora_inicio'
                elif "hora_inicio" in event_data and event_data["hora_inicio"]:
                    if (
                        isinstance(event_data["hora_inicio"], str)
                        and ":" in event_data["hora_inicio"]
                    ):
                        hora_inicio, minuto_inicio = event_data["hora_inicio"].split(
                            ":"
                        )
                    elif isinstance(event_data["hora_inicio"], datetime):
                        hora_inicio = str(event_data["hora_inicio"].hour).zfill(2)
                        minuto_inicio = str(event_data["hora_inicio"].minute).zfill(2)
                # Caso contrário, usar o valor padrão já definido

                # Primeiro tentar 'fim'
                if "fim" in event_data and event_data["fim"]:
                    if isinstance(event_data["fim"], str) and ":" in event_data["fim"]:
                        hora_fim, minuto_fim = event_data["fim"].split(":")
                    elif isinstance(event_data["fim"], datetime):
                        hora_fim = str(event_data["fim"].hour).zfill(2)
                        minuto_fim = str(event_data["fim"].minute).zfill(2)
                # Depois tentar 'hora_fim'
                elif "hora_fim" in event_data and event_data["hora_fim"]:
                    if (
                        isinstance(event_data["hora_fim"], str)
                        and ":" in event_data["hora_fim"]
                    ):
                        hora_fim, minuto_fim = event_data["hora_fim"].split(":")
                    elif isinstance(event_data["hora_fim"], datetime):
                        hora_fim = str(event_data["hora_fim"].hour).zfill(2)
                        minuto_fim = str(event_data["hora_fim"].minute).zfill(2)
                # Caso contrário, usar o valor padrão já definido

            # Selecionando o horário de inicio
            input_horario_inicio_horas = self.driver.find_element(
                By.XPATH, "//input[@name='start[hour]']"
            )
            input_horario_inicio_horas.click()
            input_horario_inicio_horas.clear()
            input_horario_inicio_horas.send_keys(hora_inicio)

            input_horario_inicio_minutos = self.driver.find_element(
                By.XPATH, "//input[@name='start[min]']"
            )
            input_horario_inicio_minutos.click()
            input_horario_inicio_minutos.clear()
            input_horario_inicio_minutos.send_keys(minuto_inicio)

            # Selecionando o horário de fim
            input_horario_fim_hora = self.driver.find_element(
                By.XPATH, "//input[@name='end[hour]']"
            )
            input_horario_fim_hora.click()
            input_horario_fim_hora.clear()
            input_horario_fim_hora.send_keys(hora_fim)

            input_horario_fim_minutos = self.driver.find_element(
                By.XPATH, "//input[@name='end[min]']"
            )
            input_horario_fim_minutos.click()
            input_horario_fim_minutos.clear()
            input_horario_fim_minutos.send_keys(minuto_fim)  # Primeiro envia os minutos
            """ input_horario_fim_minutos.send_keys(Keys.RETURN)  # Depois envia o ENTER """

            input_submit_salvar = self.driver.find_element(By.ID, "submit_button")
            input_submit_salvar.click()
            time.sleep(10)

            # Tentar obter o ID do evento criado da URL
            current_url = self.driver.current_url
            id_evento = ""

            try:
                parsed_url = urlparse(current_url)
                params = parse_qs(parsed_url.query)
                id_evento = params.get("cal_id", [""])[0] if "cal_id" in params else ""
            except:
                # Se não conseguir obter o ID, buscar o evento na lista
                eventos = self.obterEventos()
                for evento in eventos:
                    if evento.get("titulo") == event_data.get("titulo") and evento.get(
                        "data"
                    ) == event_data.get("data"):
                        id_evento = evento.get("id")
                        break

            # Adicionar ID ao evento
            event_data["id"] = id_evento

            print(f"Evento criado no Expresso com ID: {id_evento}")
            return event_data

        except Exception as e:
            print(f"Erro ao criar evento no Expresso: {e}")
            raise e

    def update_event(self, event_id, event_data):
        try:
            # Garantir que estamos na página de calendário
            if not self.driver.current_url.startswith(
                "https://www.expresso.pe.gov.br/calendar/"
            ):
                self.selecionarCalendario()

            # Navegar para a página de edição do evento
            edit_url = f"https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.edit&cal_id={event_id}"
            if "data" in event_data and event_data["data"]:
                edit_url += f"&date={event_data['data']}"

            self.driver.get(edit_url)
            time.sleep(2)

            # Atualizar os campos
            try:
                # Título
                if "titulo" in event_data and event_data["titulo"]:
                    titulo_input = self.driver.find_element(By.NAME, "title")
                    titulo_input.clear()
                    titulo_input.send_keys(event_data["titulo"])

                # Descrição
                if "descricao" in event_data and event_data["descricao"]:
                    descricao_input = self.driver.find_element(By.NAME, "description")
                    descricao_input.clear()
                    descricao_input.send_keys(event_data["descricao"])

                # Data
                if "data" in event_data and event_data["data"]:
                    data_input = self.driver.find_element(By.NAME, "date")
                    data_input.clear()
                    data_input.send_keys(event_data["data"])

                # Hora de início
                if "hora_inicio" in event_data and event_data["hora_inicio"]:
                    if isinstance(event_data["hora_inicio"], str):
                        hora_inicio = event_data["hora_inicio"].split(":")
                    else:
                        hora_inicio = [
                            str(event_data["hora_inicio"].hour),
                            str(event_data["hora_inicio"].minute),
                        ]

                    hora_input = self.driver.find_element(By.NAME, "hour")
                    hora_input.clear()
                    hora_input.send_keys(hora_inicio[0])

                    minuto_input = self.driver.find_element(By.NAME, "minute")
                    minuto_input.clear()
                    minuto_input.send_keys(hora_inicio[1])

                # Hora de término
                if "hora_fim" in event_data and event_data["hora_fim"]:
                    if isinstance(event_data["hora_fim"], str):
                        hora_fim = event_data["hora_fim"].split(":")
                    else:
                        hora_fim = [
                            str(event_data["hora_fim"].hour),
                            str(event_data["hora_fim"].minute),
                        ]

                    hora_fim_input = self.driver.find_element(By.NAME, "endhour")
                    hora_fim_input.clear()
                    hora_fim_input.send_keys(hora_fim[0])

                    minuto_fim_input = self.driver.find_element(By.NAME, "endminute")
                    minuto_fim_input.clear()
                    minuto_fim_input.send_keys(hora_fim[1])

                # Participantes
                """ if 'participantes' in event_data and event_data['participantes']:
                    participantes_input = self.driver.find_element(By.NAME, 'participants')
                    participantes_input.clear()
                    participantes_input.send_keys(event_data['participantes']) """

                # Aqui que eu comentei o código para salvar o evento
                # Salvar o evento
                salvar_button = self.driver.find_element(
                    By.XPATH, "//input[@id='submit_button']"
                )
                salvar_button.click()
                time.sleep(10)

                print(f"Evento atualizado no Expresso: {event_id}")
                return True

            except Exception as e:
                print(f"Erro ao preencher campos para atualização do evento: {e}")
                raise e

        except Exception as e:
            print(f"Erro ao atualizar evento no Expresso: {e}")
            raise e

    def delete_event(self, event_id, event_data=None):
        try:
            # Se event_data não for fornecido, usar um dicionário vazio
            if event_data is None:
                event_data = {"data": ""}

            # Garantir que estamos na página de calendário
            if not self.driver.current_url.startswith(
                "https://www.expresso.pe.gov.br/calendar/"
            ):
                self.selecionarCalendario()

            # Encontrar o evento pelo ID
            view_url = f"https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.view&cal_id={event_id}"
            if event_data.get("data"):
                view_url += f"&date={event_data['data']}"

            self.driver.get(view_url)
            time.sleep(10)

            # Clicar no botão de deletar
            try:
                botao_deletar = self.driver.find_element(
                    By.XPATH, "//input[@value='remover']"
                )
                botao_deletar.click()
                time.sleep(10)

                # Aceitar o alerta de confirmação
                alert = self.driver.switch_to.alert
                alert.accept()
                time.sleep(10)
            except:
                # Se não encontrar o botão ou não houver alerta, tentar pressionar Enter
                try:
                    # Tentar outros possíveis botões
                    botao_deletar = self.driver.find_element(
                        By.XPATH, "//input[@value='Remover']"
                    )
                    botao_deletar.click()
                    time.sleep(10)

                    # Tentar aceitar alerta, se houver
                    try:
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                except:
                    # Se ainda não conseguir, pressionar Enter
                    self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
                    time.sleep(10)

            print(f"Evento deletado no Expresso: {event_id}")
            return True

        except Exception as e:
            print(f"Erro ao deletar evento no Expresso: {e}")
            raise e

    def _format_google_to_expresso(self, google_event):
        """Converte um evento do Google Calendar para o formato do Expresso"""
        expresso_event = {}

        # Título do evento
        if "summary" in google_event:
            expresso_event["titulo"] = google_event["summary"]

        # Descrição do evento
        if "description" in google_event:
            expresso_event["descricao"] = google_event["description"]

        # Data e hora
        if "start" in google_event:
            if "dateTime" in google_event["start"]:
                # Evento com horário específico
                start_dt = datetime.fromisoformat(
                    google_event["start"]["dateTime"].replace("Z", "+00:00")
                )

                # Convertendo para o fuso horário local se necessário
                start_dt = start_dt.astimezone(tz=None)

                # Extraindo data no formato DD/MM/YYYY
                expresso_event["data"] = start_dt.strftime("%d/%m/%Y")

                # Extraindo hora de início
                expresso_event["hora_inicio"] = start_dt
            elif "date" in google_event["start"]:
                # Evento de dia inteiro
                date_obj = datetime.fromisoformat(google_event["start"]["date"])
                expresso_event["data"] = date_obj.strftime("%d/%m/%Y")
                # Definir horário padrão para eventos de dia inteiro
                expresso_event["hora_inicio"] = "00:00"

        # Hora de término
        if "end" in google_event:
            if "dateTime" in google_event["end"]:
                end_dt = datetime.fromisoformat(
                    google_event["end"]["dateTime"].replace("Z", "+00:00")
                )
                end_dt = end_dt.astimezone(tz=None)
                expresso_event["hora_fim"] = end_dt
            elif "date" in google_event["end"]:
                # Para eventos de dia inteiro, definir o final do dia
                expresso_event["hora_fim"] = "23:59"

        # Participantes
        if "attendees" in google_event:
            participantes = []
            for attendee in google_event["attendees"]:
                if "email" in attendee:
                    participantes.append(attendee["email"])
            expresso_event["participantes"] = ", ".join(participantes)

        # Localização
        if "location" in google_event:
            expresso_event["localizacao"] = google_event["location"]

        # ID do evento original para referência
        if "id" in google_event:
            expresso_event["google_id"] = google_event["id"]

        return expresso_event

    def _format_outlook_to_expresso(self, outlook_event):
        """Converte um evento do Outlook para o formato do Expresso"""
        expresso_event = {}

        # Título do evento
        if "subject" in outlook_event:
            expresso_event["titulo"] = outlook_event["subject"]

        # Descrição do evento
        if "bodyPreview" in outlook_event:
            expresso_event["descricao"] = outlook_event["bodyPreview"]
        elif "body" in outlook_event and "content" in outlook_event["body"]:
            expresso_event["descricao"] = outlook_event["body"]["content"]

        # Data e hora
        if "start" in outlook_event and "dateTime" in outlook_event["start"]:
            start_dt = datetime.fromisoformat(
                outlook_event["start"]["dateTime"].replace("Z", "+00:00")
            )
            start_dt = start_dt.astimezone(tz=None)

            # Extraindo data no formato DD/MM/YYYY
            expresso_event["data"] = start_dt.strftime("%d/%m/%Y")

            # Extraindo hora de início
            expresso_event["hora_inicio"] = start_dt

        # Hora de término
        if "end" in outlook_event and "dateTime" in outlook_event["end"]:
            end_dt = datetime.fromisoformat(
                outlook_event["end"]["dateTime"].replace("Z", "+00:00")
            )
            end_dt = end_dt.astimezone(tz=None)
            expresso_event["hora_fim"] = end_dt

        # Participantes
        if "attendees" in outlook_event:
            participantes = []
            for attendee in outlook_event["attendees"]:
                if "emailAddress" in attendee and "address" in attendee["emailAddress"]:
                    participantes.append(attendee["emailAddress"]["address"])
            expresso_event["participantes"] = ", ".join(participantes)

        # Localização
        if "location" in outlook_event and "displayName" in outlook_event["location"]:
            expresso_event["localizacao"] = outlook_event["location"]["displayName"]

        # ID do evento original para referência
        if "id" in outlook_event:
            expresso_event["outlook_id"] = outlook_event["id"]

        return expresso_event

    def _format_expresso_to_google(self, expresso_event):
        """Converte um evento do Expresso para o formato do Google Calendar"""
        google_event = {}

        # Título do evento
        if "titulo" in expresso_event:
            google_event["summary"] = expresso_event["titulo"]

        # Descrição do evento
        if "descricao" in expresso_event:
            google_event["description"] = expresso_event["descricao"]

        # Data e hora
        if "data" in expresso_event:
            data_str = expresso_event["data"]

            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if "/" in data_str:
                dia, mes, ano = data_str.split("/")
                # Garantir que os valores tenham dois dígitos
                data_iso = f"{ano}-{mes.zfill(2)}-{dia.zfill(2)}"
                print(f"Data convertida: {data_str} -> {data_iso}")  # Log para debug
            else:
                data_iso = data_str

            # Verificar se é evento de dia inteiro ou com horário específico
            dia_inteiro = expresso_event.get("dia_inteiro", False)
            
            if dia_inteiro:
                # Evento de dia inteiro
                google_event["start"] = {"date": data_iso}
                # A data de término deve ser o dia seguinte para eventos de dia inteiro
                end_date = (datetime.fromisoformat(data_iso) + timedelta(days=1)).date().isoformat()
                google_event["end"] = {"date": end_date}
            else:
                # Evento com horário específico
                if "inicio" in expresso_event and expresso_event["inicio"]:
                    hora_inicio = expresso_event["inicio"]
                    if isinstance(hora_inicio, str) and ":" in hora_inicio:
                        hora, minuto = hora_inicio.split(":")
                        # Garantir que hora e minuto tenham dois dígitos
                        start_iso = f"{data_iso}T{hora.zfill(2)}:{minuto.zfill(2)}:00"
                        google_event["start"] = {
                            "dateTime": start_iso,
                            "timeZone": "America/Recife"
                        }
                    elif isinstance(hora_inicio, datetime):
                        start_iso = hora_inicio.strftime("%Y-%m-%dT%H:%M:%S")
                        google_event["start"] = {
                            "dateTime": start_iso,
                            "timeZone": "America/Recife"
                        }

                # Hora de término
                if "fim" in expresso_event and expresso_event["fim"]:
                    hora_fim = expresso_event["fim"]
                    if isinstance(hora_fim, str) and ":" in hora_fim:
                        hora, minuto = hora_fim.split(":")
                        # Garantir que hora e minuto tenham dois dígitos
                        end_iso = f"{data_iso}T{hora.zfill(2)}:{minuto.zfill(2)}:00"
                        google_event["end"] = {
                            "dateTime": end_iso,
                            "timeZone": "America/Recife"
                        }
                    elif isinstance(hora_fim, datetime):
                        end_iso = hora_fim.strftime("%Y-%m-%dT%H:%M:%S")
                        google_event["end"] = {
                            "dateTime": end_iso,
                            "timeZone": "America/Recife"
                        }

        # Participantes com validação de e-mail
        if "participantes" in expresso_event and expresso_event["participantes"]:
            valid_attendees = []
            for email in expresso_event["participantes"].split(","):
                email = email.strip()
                # Validação básica de e-mail
                if "@" in email and "." in email.split("@")[1]:
                    valid_attendees.append({"email": email})
                else:
                    print(f"E-mail inválido ignorado: {email}")
            
            if valid_attendees:
                google_event["attendees"] = valid_attendees

        # Localização
        if "localizacao" in expresso_event:
            google_event["location"] = expresso_event["localizacao"]

        # ID do evento original para referência
        if "id" in expresso_event:
            if "extendedProperties" not in google_event:
                google_event["extendedProperties"] = {"private": {}}
            google_event["extendedProperties"]["private"]["expresso_id"] = expresso_event["id"]

        return google_event

    def _format_expresso_to_outlook(self, expresso_event):
        """Converte um evento do Expresso para o formato do Outlook"""
        outlook_event = {}

        # Título do evento
        if "titulo" in expresso_event:
            outlook_event["subject"] = expresso_event["titulo"]

        # Descrição do evento
        if "descricao" in expresso_event:
            outlook_event["body"] = {
                "contentType": "text",
                "content": expresso_event["descricao"],
            }

        # Data e hora
        if "data" in expresso_event:
            data_str = expresso_event["data"]

            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if "/" in data_str:
                dia, mes, ano = data_str.split("/")
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str

            if "hora_inicio" in expresso_event:
                # Evento com horário específico
                if isinstance(expresso_event["hora_inicio"], datetime):
                    hora_inicio = expresso_event["hora_inicio"]
                    start_iso = hora_inicio.isoformat()
                else:
                    # Assumindo formato HH:MM ou objeto de hora
                    hora, minuto = (
                        expresso_event["hora_inicio"].split(":")
                        if isinstance(expresso_event["hora_inicio"], str)
                        else [0, 0]
                    )
                    start_iso = f"{data_iso}T{hora}:{minuto}:00"

                outlook_event["start"] = {
                    "dateTime": start_iso,
                    "timeZone": "America/Recife",
                }
            else:
                # Evento de dia inteiro
                outlook_event["isAllDay"] = True
                outlook_event["start"] = {
                    "dateTime": f"{data_iso}T00:00:00",
                    "timeZone": "America/Recife",
                }

        # Hora de término
        if "data" in expresso_event and "hora_fim" in expresso_event:
            data_str = expresso_event["data"]

            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if "/" in data_str:
                dia, mes, ano = data_str.split("/")
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str

            if isinstance(expresso_event["hora_fim"], datetime):
                hora_fim = expresso_event["hora_fim"]
                end_iso = hora_fim.isoformat()
            else:
                # Assumindo formato HH:MM ou objeto de hora
                hora, minuto = (
                    expresso_event["hora_fim"].split(":")
                    if isinstance(expresso_event["hora_fim"], str)
                    else [23, 59]
                )
                end_iso = f"{data_iso}T{hora}:{minuto}:00"

            outlook_event["end"] = {"dateTime": end_iso, "timeZone": "America/Recife"}
        elif "isAllDay" in outlook_event and outlook_event["isAllDay"]:
            # Para eventos de dia inteiro, definir o final do dia
            data_str = expresso_event["data"]
            if "/" in data_str:
                dia, mes, ano = data_str.split("/")
                # Para eventos de dia inteiro, a data de término deve ser o dia seguinte às 00:00
                # Criar objeto de data e adicionar um dia
                from datetime import datetime, timedelta

                data_obj = datetime(int(ano), int(mes), int(dia))
                data_seguinte = data_obj + timedelta(days=1)
                data_iso_seguinte = data_seguinte.strftime("%Y-%m-%d")

                outlook_event["end"] = {
                    "dateTime": f"{data_iso_seguinte}T00:00:00",
                    "timeZone": "America/Recife",
                }
            else:
                # Se não conseguir parsear a data, tentar usar a data original + 1 dia
                try:
                    from datetime import datetime, timedelta

                    data_obj = datetime.fromisoformat(data_iso)
                    data_seguinte = data_obj + timedelta(days=1)
                    data_iso_seguinte = data_seguinte.strftime("%Y-%m-%d")

                    outlook_event["end"] = {
                        "dateTime": f"{data_iso_seguinte}T00:00:00",
                        "timeZone": "America/Recife",
                    }
                except:
                    # Fallback se não conseguir calcular a data seguinte
                    outlook_event["end"] = {
                        "dateTime": f"{data_iso}T00:00:00",
                        "timeZone": "America/Recife",
                    }

        # Participantes
        if "participantes" in expresso_event and expresso_event["participantes"]:
            outlook_event["attendees"] = []
            for email in expresso_event["participantes"].split(","):
                outlook_event["attendees"].append(
                    {
                        "emailAddress": {
                            "address": email.strip(),
                            "name": email.strip(),
                        },
                        "type": "required",
                    }
                )

        # Localização
        if "localizacao" in expresso_event:
            outlook_event["location"] = {"displayName": expresso_event["localizacao"]}

        # ID do evento original para referência
        if "id" in expresso_event:
            # Outlook não tem um campo direto para armazenar IDs externos
            # Para isso, poderia ser usado um campo de extensão ou custom properties
            pass

        return outlook_event

    def _is_event_duplicate(self, event, source_type, target_events_cache):
        """Verifica se um evento já existe no cache de destino"""
        # Normalizar o evento para comparação
        normalized_event = self._normalize_event_for_comparison(event, source_type)

        # Verificar com cada evento no cache de destino
        for target_id, target_event in target_events_cache.items():
            target_normalized = self._normalize_event_for_comparison(
                target_event, "google" if source_type == "outlook" else "outlook"
            )

            # Comparar título e data
            if normalized_event["title"] == target_normalized["title"]:
                # Se os títulos são iguais, retornar o ID do evento de destino
                return True, target_id

        return False, None

    def _is_expresso_event_updated(self, current_event, cached_event):
        """Verifica se um evento do Expresso foi atualizado"""
        # Campos importantes
        fields = ['titulo', 'descricao', 'data', 'inicio', 'fim', 'localizacao']
        
        for field in fields:
            if field in current_event and field in cached_event:
                if current_event[field] != cached_event[field]:
                    return True
        
        return False

    def sync_changes_only(self):
        print("\n=== INICIANDO SINCRONIZAÇÃO ===")
        print(f"Data/Hora atual: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"Última sincronização: {self.last_sync_time.strftime('%d/%m/%Y %H:%M:%S')}")

        if hasattr(self, "expresso_sync") and self.expresso_sync:
            print("\n=== VERIFICANDO EXPRESSO ===")
            try:
                expresso_events = self.expresso_sync.obterEventos()
                print(f"Eventos obtidos do Expresso: {len(expresso_events)}")
                for event in expresso_events[:5]:  # Mostrar primeiros 5 eventos
                    print(f"  - {event.get('titulo', 'Sem título')} ({event.get('data', 'sem data')})")
            except Exception as e:
                print(f"ERRO ao obter eventos do Expresso: {e}")


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
