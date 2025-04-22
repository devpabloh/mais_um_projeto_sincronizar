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
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
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
        time.sleep(5)
        
        try:
            # Esperar até que os elementos de eventos estejam presentes
            wait = WebDriverWait(self.driver, 30)
            
            # Tentar encontrar os eventos com diferentes seletores
            try:
                eventos = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@id='calendar_event_entry']/a[@class='event_entry']")))
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
                
                # Extrair horário - versão mais robusta
                try:
                    # Tentar várias abordagens para encontrar o horário
                    # Abordagem 1: Buscar diretamente por span com cor preta
                    try:
                        span_horario = evento.find_element(By.CSS_SELECTOR, "span[style*='color: black']")
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
                                padrao_hora = re.search(r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})', texto_completo)
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
                if 'horario_texto' in locals() and '-' in horario_texto:
                    horarios = horario_texto.split('-')
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
                if (evento.get('titulo', '') == event_data.get('titulo', '') and 
                    evento.get('data', '') == event_data.get('data', '')):
                    # Verificar horário com tolerância de 5 minutos
                    hora_evento = evento.get('inicio', '').split(':')
                    hora_novo = event_data.get('inicio', '').split(':')
                    
                    if len(hora_evento) == 2 and len(hora_novo) == 2:
                        minutos_evento = int(hora_evento[0]) * 60 + int(hora_evento[1])
                        minutos_novo = int(hora_novo[0]) * 60 + int(hora_novo[1])
                        
                        if abs(minutos_evento - minutos_novo) <= 5:
                            print(f"Evento já existe no Expresso: {evento['titulo']} em {evento['data']} {evento['inicio']}")
                            # Retornar o evento existente em vez de criar um novo
                            return evento
            
            # Se não encontrou evento duplicado, continuar com a criação
            # Navegando até a página de criação de eventos
            self.driver.get("https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.add&date=20250423")
            time.sleep(10)
            
            # Preenchendo o formulário
            input_titulo = self.driver.find_element(By.XPATH, "//input[@name='cal[title]']")
            input_titulo.clear()
            input_titulo.send_keys(event_data["titulo"]) 

            # Selecionando a descrição
            input_descricao = self.driver.find_element(By.XPATH, "//textarea[@name='cal[description]']")
            input_descricao.clear()
            input_descricao.send_keys(event_data["descricao"])

            # Selecionando a localização
            input_de_licalizacao = self.driver.find_element(By.XPATH, "//input[@name='cal[location]']")
            input_de_licalizacao.clear()
            input_de_licalizacao.send_keys(event_data.get("localizacao", ""))
            
            # Selecionando a data de inicio
            input_data = self.driver.find_element(By.XPATH, "//input[@name='start[str]']")
            input_data.clear()
            input_data.send_keys(event_data["data"]) 

            # Selecionando a data de fim
            input_data_fim = self.driver.find_element(By.XPATH, "//input[@name='end[str]']")
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
                # Para eventos com horário específico
                horario_inicio = event_data.get("inicio", "00:00")
                if isinstance(horario_inicio, str) and ":" in horario_inicio:
                    hora_inicio, minuto_inicio = horario_inicio.split(":")
                elif isinstance(horario_inicio, datetime):
                    hora_inicio = str(horario_inicio.hour).zfill(2)
                    minuto_inicio = str(horario_inicio.minute).zfill(2)
                
                horario_fim = event_data.get("fim", "23:59")
                if isinstance(horario_fim, str) and ":" in horario_fim:
                    hora_fim, minuto_fim = horario_fim.split(":")
                elif isinstance(horario_fim, datetime):
                    hora_fim = str(horario_fim.hour).zfill(2)
                    minuto_fim = str(horario_fim.minute).zfill(2)

            # Selecionando o horário de inicio
            input_horario_inicio_horas = self.driver.find_element(By.XPATH, "//input[@name='start[hour]']")
            input_horario_inicio_horas.click()
            input_horario_inicio_horas.clear()
            input_horario_inicio_horas.send_keys(hora_inicio)

            input_horario_inicio_minutos = self.driver.find_element(By.XPATH, "//input[@name='start[min]']")
            input_horario_inicio_minutos.click()
            input_horario_inicio_minutos.clear()
            input_horario_inicio_minutos.send_keys(minuto_inicio)

            # Selecionando o horário de fim
            input_horario_fim_hora = self.driver.find_element(By.XPATH, "//input[@name='end[hour]']")
            input_horario_fim_hora.click()
            input_horario_fim_hora.clear()
            input_horario_fim_hora.send_keys(hora_fim)

            input_horario_fim_minutos = self.driver.find_element(By.XPATH, "//input[@name='end[min]']")
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
                    if (evento.get('titulo') == event_data.get('titulo') and 
                        evento.get('data') == event_data.get('data')):
                        id_evento = evento.get('id')
                        break
            
            # Adicionar ID ao evento
            event_data['id'] = id_evento
            
            print(f"Evento criado no Expresso com ID: {id_evento}")
            return event_data
                
        except Exception as e:
            print(f"Erro ao criar evento no Expresso: {e}")
            raise e
        
    def update_event(self, event_id, event_data):
        try:
            # Garantir que estamos na página de calendário
            if not self.driver.current_url.startswith('https://www.expresso.pe.gov.br/calendar/'):
                self.selecionarCalendario()
            
            # Navegar para a página de edição do evento
            edit_url = f"https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.edit&cal_id={event_id}"
            if 'data' in event_data and event_data['data']:
                edit_url += f"&date={event_data['data']}"
            
            self.driver.get(edit_url)
            time.sleep(2)
            
            # Atualizar os campos
            try:
                # Título
                if 'titulo' in event_data and event_data['titulo']:
                    titulo_input = self.driver.find_element(By.NAME, 'title')
                    titulo_input.clear()
                    titulo_input.send_keys(event_data['titulo'])
                
                # Descrição
                if 'descricao' in event_data and event_data['descricao']:
                    descricao_input = self.driver.find_element(By.NAME, 'description')
                    descricao_input.clear()
                    descricao_input.send_keys(event_data['descricao'])
                
                # Data
                if 'data' in event_data and event_data['data']:
                    data_input = self.driver.find_element(By.NAME, 'date')
                    data_input.clear()
                    data_input.send_keys(event_data['data'])
                
                # Hora de início
                if 'hora_inicio' in event_data and event_data['hora_inicio']:
                    if isinstance(event_data['hora_inicio'], str):
                        hora_inicio = event_data['hora_inicio'].split(':')
                    else:
                        hora_inicio = [str(event_data['hora_inicio'].hour), str(event_data['hora_inicio'].minute)]
                    
                    hora_input = self.driver.find_element(By.NAME, 'hour')
                    hora_input.clear()
                    hora_input.send_keys(hora_inicio[0])
                    
                    minuto_input = self.driver.find_element(By.NAME, 'minute')
                    minuto_input.clear()
                    minuto_input.send_keys(hora_inicio[1])
                
                # Hora de término
                if 'hora_fim' in event_data and event_data['hora_fim']:
                    if isinstance(event_data['hora_fim'], str):
                        hora_fim = event_data['hora_fim'].split(':')
                    else:
                        hora_fim = [str(event_data['hora_fim'].hour), str(event_data['hora_fim'].minute)]
                    
                    hora_fim_input = self.driver.find_element(By.NAME, 'endhour')
                    hora_fim_input.clear()
                    hora_fim_input.send_keys(hora_fim[0])
                    
                    minuto_fim_input = self.driver.find_element(By.NAME, 'endminute')
                    minuto_fim_input.clear()
                    minuto_fim_input.send_keys(hora_fim[1])
                
                # Participantes
                """ if 'participantes' in event_data and event_data['participantes']:
                    participantes_input = self.driver.find_element(By.NAME, 'participants')
                    participantes_input.clear()
                    participantes_input.send_keys(event_data['participantes']) """
                
                #Aqui que eu comentei o código para salvar o evento
                # Salvar o evento
                salvar_button = self.driver.find_element(By.XPATH, "//input[@id='submit_button']")
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
                event_data = {'data': ''}
            
            # Garantir que estamos na página de calendário
            if not self.driver.current_url.startswith('https://www.expresso.pe.gov.br/calendar/'):
                self.selecionarCalendario()

            # Encontrar o evento pelo ID
            view_url = f"https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.view&cal_id={event_id}"
            if event_data.get('data'):
                view_url += f"&date={event_data['data']}"
            
            self.driver.get(view_url)
            time.sleep(10)
            
            # Clicar no botão de deletar
            try:
                botao_deletar = self.driver.find_element(By.XPATH, "//input[@value='remover']")
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
                    botao_deletar = self.driver.find_element(By.XPATH, "//input[@value='Remover']")
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
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ENTER)
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
        if 'summary' in google_event:
            expresso_event['titulo'] = google_event['summary']
        
        # Descrição do evento
        if 'description' in google_event:
            expresso_event['descricao'] = google_event['description']
        
        # Data e hora
        if 'start' in google_event:
            if 'dateTime' in google_event['start']:
                # Evento com horário específico
                start_dt = datetime.fromisoformat(google_event['start']['dateTime'].replace('Z', '+00:00'))
                
                # Convertendo para o fuso horário local se necessário
                start_dt = start_dt.astimezone(tz=None)
                
                # Extraindo data no formato DD/MM/YYYY
                expresso_event['data'] = start_dt.strftime('%d/%m/%Y')
                
                # Extraindo hora de início
                expresso_event['hora_inicio'] = start_dt
            elif 'date' in google_event['start']:
                # Evento de dia inteiro
                date_obj = datetime.fromisoformat(google_event['start']['date'])
                expresso_event['data'] = date_obj.strftime('%d/%m/%Y')
                # Definir horário padrão para eventos de dia inteiro
                expresso_event['hora_inicio'] = '00:00'
        
        # Hora de término
        if 'end' in google_event:
            if 'dateTime' in google_event['end']:
                end_dt = datetime.fromisoformat(google_event['end']['dateTime'].replace('Z', '+00:00'))
                end_dt = end_dt.astimezone(tz=None)
                expresso_event['hora_fim'] = end_dt
            elif 'date' in google_event['end']:
                # Para eventos de dia inteiro, definir o final do dia
                expresso_event['hora_fim'] = '23:59'
        
        # Participantes
        if 'attendees' in google_event:
            participantes = []
            for attendee in google_event['attendees']:
                if 'email' in attendee:
                    participantes.append(attendee['email'])
            expresso_event['participantes'] = ', '.join(participantes)
        
        # Localização
        if 'location' in google_event:
            expresso_event['localizacao'] = google_event['location']
        
        # ID do evento original para referência
        if 'id' in google_event:
            expresso_event['google_id'] = google_event['id']
        
        return expresso_event

    def _format_outlook_to_expresso(self, outlook_event):
        """Converte um evento do Outlook para o formato do Expresso"""
        expresso_event = {}
        
        # Título do evento
        if 'subject' in outlook_event:
            expresso_event['titulo'] = outlook_event['subject']
        
        # Descrição do evento
        if 'bodyPreview' in outlook_event:
            expresso_event['descricao'] = outlook_event['bodyPreview']
        elif 'body' in outlook_event and 'content' in outlook_event['body']:
            expresso_event['descricao'] = outlook_event['body']['content']
        
        # Data e hora
        if 'start' in outlook_event and 'dateTime' in outlook_event['start']:
            start_dt = datetime.fromisoformat(outlook_event['start']['dateTime'].replace('Z', '+00:00'))
            start_dt = start_dt.astimezone(tz=None)
            
            # Extraindo data no formato DD/MM/YYYY
            expresso_event['data'] = start_dt.strftime('%d/%m/%Y')
            
            # Extraindo hora de início
            expresso_event['hora_inicio'] = start_dt
        
        # Hora de término
        if 'end' in outlook_event and 'dateTime' in outlook_event['end']:
            end_dt = datetime.fromisoformat(outlook_event['end']['dateTime'].replace('Z', '+00:00'))
            end_dt = end_dt.astimezone(tz=None)
            expresso_event['hora_fim'] = end_dt
        
        # Participantes
        if 'attendees' in outlook_event:
            participantes = []
            for attendee in outlook_event['attendees']:
                if 'emailAddress' in attendee and 'address' in attendee['emailAddress']:
                    participantes.append(attendee['emailAddress']['address'])
            expresso_event['participantes'] = ', '.join(participantes)
        
        # Localização
        if 'location' in outlook_event and 'displayName' in outlook_event['location']:
            expresso_event['localizacao'] = outlook_event['location']['displayName']
        
        # ID do evento original para referência
        if 'id' in outlook_event:
            expresso_event['outlook_id'] = outlook_event['id']
        
        return expresso_event

    def _format_expresso_to_google(self, expresso_event):
        """Converte um evento do Expresso para o formato do Google Calendar"""
        google_event = {}
        
        # Título do evento
        if 'titulo' in expresso_event:
            google_event['summary'] = expresso_event['titulo']
        
        # Descrição do evento
        if 'descricao' in expresso_event:
            google_event['description'] = expresso_event['descricao']
        
        # Data e hora
        if 'data' in expresso_event:
            data_str = expresso_event['data']
            
            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if '/' in data_str:
                dia, mes, ano = data_str.split('/')
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str
            
            if 'hora_inicio' in expresso_event:
                # Evento com horário específico
                if isinstance(expresso_event['hora_inicio'], datetime):
                    hora_inicio = expresso_event['hora_inicio']
                    start_iso = hora_inicio.isoformat()
                else:
                    # Assumindo formato HH:MM ou objeto de hora
                    hora, minuto = expresso_event['hora_inicio'].split(':') if isinstance(expresso_event['hora_inicio'], str) else [0, 0]
                    start_iso = f"{data_iso}T{hora}:{minuto}:00"
                
                google_event['start'] = {'dateTime': start_iso, 'timeZone': 'America/Recife'}
            else:
                # Evento de dia inteiro
                google_event['start'] = {'date': data_iso}
        
        # Hora de término
        if 'data' in expresso_event and 'hora_fim' in expresso_event:
            data_str = expresso_event['data']
            
            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if '/' in data_str:
                dia, mes, ano = data_str.split('/')
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str
            
            if isinstance(expresso_event['hora_fim'], datetime):
                hora_fim = expresso_event['hora_fim']
                end_iso = hora_fim.isoformat()
            else:
                # Assumindo formato HH:MM ou objeto de hora
                hora, minuto = expresso_event['hora_fim'].split(':') if isinstance(expresso_event['hora_fim'], str) else [23, 59]
                end_iso = f"{data_iso}T{hora}:{minuto}:00"
            
            google_event['end'] = {'dateTime': end_iso, 'timeZone': 'America/Recife'}
        elif 'start' in google_event and 'date' in google_event['start']:
            # Para eventos de dia inteiro, a data de término é o dia seguinte
            end_date = (datetime.fromisoformat(google_event['start']['date']) + timedelta(days=1)).date().isoformat()
            google_event['end'] = {'date': end_date}
        
        # Participantes
        if 'participantes' in expresso_event and expresso_event['participantes']:
            google_event['attendees'] = []
            for email in expresso_event['participantes'].split(','):
                google_event['attendees'].append({'email': email.strip()})
        
        # Localização
        if 'localizacao' in expresso_event:
            google_event['location'] = expresso_event['localizacao']
        
        # ID do evento original para referência
        if 'id' in expresso_event:
            google_event['extendedProperties'] = {
                'private': {
                    'expresso_id': expresso_event['id']
                }
            }
        
        return google_event

    def _format_expresso_to_outlook(self, expresso_event):
        """Converte um evento do Expresso para o formato do Outlook"""
        outlook_event = {}
        
        # Título do evento
        if 'titulo' in expresso_event:
            outlook_event['subject'] = expresso_event['titulo']
        
        # Descrição do evento
        if 'descricao' in expresso_event:
            outlook_event['body'] = {
                'contentType': 'text',
                'content': expresso_event['descricao']
            }
        
        # Data e hora
        if 'data' in expresso_event:
            data_str = expresso_event['data']
            
            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if '/' in data_str:
                dia, mes, ano = data_str.split('/')
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str
            
            if 'hora_inicio' in expresso_event:
                # Evento com horário específico
                if isinstance(expresso_event['hora_inicio'], datetime):
                    hora_inicio = expresso_event['hora_inicio']
                    start_iso = hora_inicio.isoformat()
                else:
                    # Assumindo formato HH:MM ou objeto de hora
                    hora, minuto = expresso_event['hora_inicio'].split(':') if isinstance(expresso_event['hora_inicio'], str) else [0, 0]
                    start_iso = f"{data_iso}T{hora}:{minuto}:00"
                
                outlook_event['start'] = {
                    'dateTime': start_iso,
                    'timeZone': 'America/Recife'
                }
            else:
                # Evento de dia inteiro
                outlook_event['isAllDay'] = True
                outlook_event['start'] = {
                    'dateTime': f"{data_iso}T00:00:00",
                    'timeZone': 'America/Recife'
                }
        
        # Hora de término
        if 'data' in expresso_event and 'hora_fim' in expresso_event:
            data_str = expresso_event['data']
            
            # Convertendo data do formato DD/MM/YYYY para YYYY-MM-DD
            if '/' in data_str:
                dia, mes, ano = data_str.split('/')
                data_iso = f"{ano}-{mes}-{dia}"
            else:
                data_iso = data_str
            
            if isinstance(expresso_event['hora_fim'], datetime):
                hora_fim = expresso_event['hora_fim']
                end_iso = hora_fim.isoformat()
            else:
                # Assumindo formato HH:MM ou objeto de hora
                hora, minuto = expresso_event['hora_fim'].split(':') if isinstance(expresso_event['hora_fim'], str) else [23, 59]
                end_iso = f"{data_iso}T{hora}:{minuto}:00"
            
            outlook_event['end'] = {
                'dateTime': end_iso,
                'timeZone': 'America/Recife'
            }
        elif 'isAllDay' in outlook_event and outlook_event['isAllDay']:
            # Para eventos de dia inteiro, definir o final do dia
            data_str = expresso_event['data']
            if '/' in data_str:
                dia, mes, ano = data_str.split('/')
                # Para eventos de dia inteiro, a data de término deve ser o dia seguinte às 00:00
                # Criar objeto de data e adicionar um dia
                from datetime import datetime, timedelta
                data_obj = datetime(int(ano), int(mes), int(dia))
                data_seguinte = data_obj + timedelta(days=1)
                data_iso_seguinte = data_seguinte.strftime('%Y-%m-%d')
                
                outlook_event['end'] = {
                    'dateTime': f"{data_iso_seguinte}T00:00:00",
                    'timeZone': 'America/Recife'
                }
            else:
                # Se não conseguir parsear a data, tentar usar a data original + 1 dia
                try:
                    from datetime import datetime, timedelta
                    data_obj = datetime.fromisoformat(data_iso)
                    data_seguinte = data_obj + timedelta(days=1)
                    data_iso_seguinte = data_seguinte.strftime('%Y-%m-%d')
                    
                    outlook_event['end'] = {
                        'dateTime': f"{data_iso_seguinte}T00:00:00",
                        'timeZone': 'America/Recife'
                    }
                except:
                    # Fallback se não conseguir calcular a data seguinte
                    outlook_event['end'] = {
                        'dateTime': f"{data_iso}T00:00:00",
                        'timeZone': 'America/Recife'
                    }
        
        # Participantes
        if 'participantes' in expresso_event and expresso_event['participantes']:
            outlook_event['attendees'] = []
            for email in expresso_event['participantes'].split(','):
                outlook_event['attendees'].append({
                    'emailAddress': {
                        'address': email.strip(),
                        'name': email.strip()
                    },
                    'type': 'required'
                })
        
        # Localização
        if 'localizacao' in expresso_event:
            outlook_event['location'] = {
                'displayName': expresso_event['localizacao']
            }
        
        # ID do evento original para referência
        if 'id' in expresso_event:
            # Outlook não tem um campo direto para armazenar IDs externos
            # Para isso, poderia ser usado um campo de extensão ou custom properties
            pass
        
        return outlook_event

    def _is_duplicate_event(self, event_data, source_type, target_type, target_events):
        """
        Verifica se um evento da fonte já existe no calendário de destino.
        
        Args:
            event_data: Os dados do evento da fonte
            source_type: 'google', 'outlook' ou 'expresso'
            target_type: 'google', 'outlook' ou 'expresso'
            target_events: Dicionário de eventos do calendário de destino
        
        Returns:
            tuple: (é_duplicado, id_evento_existente)
        """
        # Obter título e data/hora com base no tipo de origem
        if source_type == 'google':
            title = event_data.get('summary', '')
            description = event_data.get('description', '')
            start_datetime = event_data.get('start', {}).get('dateTime', '')
            try:
                start_dt = datetime.fromisoformat(start_datetime.replace('Z', '+00:00'))
            except (ValueError, TypeError):
                start_dt = None
        elif source_type == 'outlook':
            title = event_data.get('subject', '')
            description = event_data.get('body', {}).get('content', '') if isinstance(event_data.get('body'), dict) else ''
            start_datetime = event_data.get('start', {}).get('dateTime', '')
            try:
                start_dt = datetime.fromisoformat(start_datetime.replace('Z', '+00:00'))
            except (ValueError, TypeError):
                start_dt = None
        elif source_type == 'expresso':
            title = event_data.get('titulo', '')
            description = event_data.get('descricao', '')
            data = event_data.get('data', '')
            inicio = event_data.get('inicio', '')
            try:
                if data and inicio and ':' in inicio:
                    dia, mes, ano = data.split('/')
                    hora, minuto = inicio.split(':')
                    start_dt = datetime(int(ano), int(mes), int(dia), int(hora), int(minuto))
                else:
                    start_dt = None
            except (ValueError, IndexError):
                start_dt = None
        else:
            return False, None
        
        # Se não temos título ou data/hora, não podemos comparar
        if not title or not start_dt:
            return False, None
        
        print(f"Verificando se evento '{title}' ({start_dt}) já existe no {target_type}...")
        
        # Verificar cada evento no destino
        for event_id, event in target_events.items():
            # Extrair dados do evento de destino conforme seu tipo
            if target_type == 'google':
                target_title = event.get('summary', '')
                target_description = event.get('description', '')
                target_start_str = event.get('start', {}).get('dateTime', '')
                try:
                    target_start = datetime.fromisoformat(target_start_str.replace('Z', '+00:00'))
                except (ValueError, TypeError):
                    target_start = None
            elif target_type == 'outlook':
                target_title = event.get('subject', '')
                target_description = event.get('body', {}).get('content', '') if isinstance(event.get('body'), dict) else ''
                target_start_str = event.get('start', {}).get('dateTime', '')
                try:
                    target_start = datetime.fromisoformat(target_start_str.replace('Z', '+00:00'))
                except (ValueError, TypeError):
                    target_start = None
            elif target_type == 'expresso':
                target_title = event.get('titulo', '')
                target_description = event.get('descricao', '')
                data = event.get('data', '')
                inicio = event.get('inicio', '')
                try:
                    if data and inicio and ':' in inicio:
                        dia, mes, ano = data.split('/')
                        hora, minuto = inicio.split(':')
                        target_start = datetime(int(ano), int(mes), int(dia), int(hora), int(minuto))
                    else:
                        target_start = None
                except (ValueError, IndexError):
                    target_start = None
            else:
                continue
            
            # Sem título ou sem data/hora, não podemos comparar
            if not target_title or not target_start:
                continue
            
            # VERIFICAÇÃO 1: Títulos idênticos (ignorando case)
            titles_match = title.lower() == target_title.lower()
            
            # VERIFICAÇÃO 2: Verificação de proximidade de título usando distância de Levenshtein
            # Se os títulos forem muito parecidos (ex: "Reunião" vs "Reuniao"), considerar match
            title_similarity = 0
            if not titles_match and len(title) > 3 and len(target_title) > 3:
                # Calcular semelhança de strings (implementação simples)
                if len(title) > len(target_title):
                    title, target_title = target_title, title  # Garantir que title é a menor string
                
                # Calcular quantos caracteres são iguais na mesma posição
                matches = sum(1 for a, b in zip(title.lower(), target_title.lower()) if a == b)
                if len(title) > 0:
                    title_similarity = matches / len(title)
                
                # Se mais de 80% dos caracteres correspondem, considerar como match
                titles_match = title_similarity > 0.8
            
            # VERIFICAÇÃO 3: Data/hora com tolerância
            times_match = False
            if target_start and start_dt:
                time_diff = abs((start_dt - target_start).total_seconds())
                # Tolerância de 5 minutos
                times_match = time_diff <= 300
            
            # VERIFICAÇÃO 4: Similaridade de descrição (se ambos tiverem descrição)
            desc_match = False
            if description and target_description:
                # Limpeza básica para comparação
                desc1 = description.lower().replace('\n', ' ').replace('\r', '').strip()
                desc2 = target_description.lower().replace('\n', ' ').replace('\r', '').strip()
                
                # Se uma descrição contém a outra completamente
                if desc1 in desc2 or desc2 in desc1:
                    desc_match = True
                
                # Ou se elas compartilham palavras-chave significativas (simples)
                elif len(desc1) > 10 and len(desc2) > 10:
                    words1 = set(w for w in desc1.split() if len(w) > 4)  # Palavras com mais de 4 letras
                    words2 = set(w for w in desc2.split() if len(w) > 4)
                    if words1 and words2:
                        common_words = words1.intersection(words2)
                        if len(common_words) >= 3:  # Se compartilham pelo menos 3 palavras significativas
                            desc_match = True
            
            # Para considerar um evento como duplicado:
            # 1. Título e horário devem corresponder 
            # 2. OU título e descrição devem corresponder significativamente
            if (titles_match and times_match) or (titles_match and desc_match):
                print(f"  - Evento duplicado encontrado no {target_type}: {target_title} ({target_start})")
                print(f"  - Similaridade: títulos={titles_match}, horários={times_match}, descrição={desc_match}")
                return True, event_id
        
        return False, None

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
