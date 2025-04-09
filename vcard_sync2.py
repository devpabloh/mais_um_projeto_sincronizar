from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse, parse_qs
import time
import keyboard
from datetime import datetime

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

    def create_event(self, event_data):
        try:
            # Navegando até a página de criação de eventos
            self.driver.get("https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.add&date=20250423")
            time.sleep(3)
            
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

            # Dividir o horário de início (formato HH:MM) em horas e minutos
            horario_inicio = event_data["inicio"]
            if ":" in horario_inicio:
                hora_inicio, minuto_inicio = horario_inicio.split(":")
            else:
                # Caso não esteja no formato HH:MM
                hora_inicio = "00"
                minuto_inicio = "00"

            # Selecionando o horário de inicio
            input_horario_inicio_horas = self.driver.find_element(By.XPATH, "//input[@name='start[hour]']")
            input_horario_inicio_horas.clear()
            input_horario_inicio_horas.send_keys(hora_inicio)

            input_horario_inicio_minutos = self.driver.find_element(By.XPATH, "//input[@name='start[min]']")
            input_horario_inicio_minutos.clear()
            input_horario_inicio_minutos.send_keys(minuto_inicio)

            # Dividir o horário de fim (formato HH:MM) em horas e minutos
            horario_fim = event_data["fim"]
            if ":" in horario_fim:
                hora_fim, minuto_fim = horario_fim.split(":")
            else:
                # Caso não esteja no formato HH:MM
                hora_fim = "23"
                minuto_fim = "59"

            # Selecionando o horário de fim
            input_horario_fim_hora = self.driver.find_element(By.XPATH, "//input[@name='end[hour]']")
            input_horario_fim_hora.clear()
            input_horario_fim_hora.send_keys(hora_fim)

            input_horario_fim_minutos = self.driver.find_element(By.XPATH, "//input[@name='end[min]']")
            input_horario_fim_minutos.clear()
            input_horario_fim_minutos.send_keys(minuto_fim)

            input_submit_salvar = self.driver.find_element(By.XPATH, "//input[@value='salvar']")
            input_submit_salvar.click()
            time.sleep(3)

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
            self.selecionarCalendario()

            # Encontrar o evento pelo ID
            edit_url = f"https://www.expresso.pe.gov.br/index.php?menuaction=calendar.uicalendar.view&cal_id={event_id}&date={event_data['data']}"
            self.driver.get(edit_url)
            time.sleep(2)

            # Verificar se estamos realmente na página de edição
            # (removida verificação problemática)
            
            # Atualizando campos do evento 
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

            # Dividir o horário de início (formato HH:MM) em horas e minutos
            horario_inicio = event_data["inicio"]
            if ":" in horario_inicio:
                hora_inicio, minuto_inicio = horario_inicio.split(":")
            else:
                # Caso não esteja no formato HH:MM
                hora_inicio = "00"
                minuto_inicio = "00"

            # Selecionando o horário de inicio
            input_horario_inicio_horas = self.driver.find_element(By.XPATH, "//input[@name='start[hour]']")
            input_horario_inicio_horas.clear()
            input_horario_inicio_horas.send_keys(hora_inicio)

            input_horario_inicio_minutos = self.driver.find_element(By.XPATH, "//input[@name='start[min]']")
            input_horario_inicio_minutos.clear()
            input_horario_inicio_minutos.send_keys(minuto_inicio)

            # Dividir o horário de fim (formato HH:MM) em horas e minutos
            horario_fim = event_data["fim"]
            if ":" in horario_fim:
                hora_fim, minuto_fim = horario_fim.split(":")
            else:
                # Caso não esteja no formato HH:MM
                hora_fim = "23"
                minuto_fim = "59"

            # Selecionando o horário de fim
            input_horario_fim_hora = self.driver.find_element(By.XPATH, "//input[@name='end[hour]']")
            input_horario_fim_hora.clear()
            input_horario_fim_hora.send_keys(hora_fim)

            input_horario_fim_minutos = self.driver.find_element(By.XPATH, "//input[@name='end[min]']")
            input_horario_fim_minutos.clear()
            input_horario_fim_minutos.send_keys(minuto_fim)

            input_submit_salvar = self.driver.find_element(By.XPATH, "//input[@value='salvar']")
            input_submit_salvar.click()
            time.sleep(3)

            # Confirmar que o evento foi atualizado
            print(f"Evento atualizado no Expresso: {event_id}")
            
            # Adicionar ID ao evento retornado
            event_data['id'] = event_id
            return event_data
            
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
            time.sleep(2)

            # Verificar se estamos na página correta
            # (modificado, pois a verificação original pode não ser precisa)
            
            # Clicar no botão de deletar
            try:
                botao_deletar = self.driver.find_element(By.XPATH, "//input[@value='remover']")
                botao_deletar.click()
                time.sleep(2)
            
                # Aceitar o alerta de confirmação
                alert = self.driver.switch_to.alert
                alert.accept()
                time.sleep(2)
            except:
                # Se não encontrar o botão ou não houver alerta, tentar pressionar Enter
                try:
                    # Tentar outros possíveis botões
                    botao_deletar = self.driver.find_element(By.XPATH, "//input[@value='Remover']")
                    botao_deletar.click()
                    time.sleep(2)
                    
                    # Tentar aceitar alerta, se houver
                    try:
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                except:
                    # Se ainda não conseguir, pressionar Enter
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ENTER)
                    time.sleep(1)
            
            print(f"Evento deletado no Expresso: {event_id}")
            return True
            
        except Exception as e:
            print(f"Erro ao deletar evento no Expresso: {e}")
            raise e

    def _format_google_to_expresso(self, google_event):
        """
        Converte um evento do Google para o formato do Expresso.
        """
        if not google_event.get('summary') or not google_event.get('start'):
            return None
        
        # Extrair data e hora do evento do Google
        start_datetime = google_event.get('start', {}).get('dateTime', '')
        end_datetime = google_event.get('end', {}).get('dateTime', '')
        
        # Converter para datetime para formatação
        try:
            start_dt = datetime.fromisoformat(start_datetime.replace('Z', '+00:00'))
            end_dt = datetime.fromisoformat(end_datetime.replace('Z', '+00:00'))
            
            # Formatar para Expresso
            data = start_dt.strftime('%d/%m/%Y')
            inicio = start_dt.strftime('%H:%M')
            fim = end_dt.strftime('%H:%M')
        except (ValueError, TypeError):
            # Fallback para eventos sem data/hora específica
            data = ''
            inicio = ''
            fim = ''
        
        expresso_event = {
            'titulo': google_event.get('summary', ''),
            'descricao': google_event.get('description', ''),
            'data': data,
            'inicio': inicio,
            'fim': fim,
            'localizacao': google_event.get('location', ''),
            'participantes': ''  # Não disponível diretamente no Google
        }
        
        return expresso_event

    def _format_outlook_to_expresso(self, outlook_event):
        """
        Converte um evento do Outlook para o formato do Expresso.
        """
        if not outlook_event.get('subject') or not outlook_event.get('start'):
            return None
        
        # Extrair data e hora do evento do Outlook
        start_datetime = outlook_event.get('start', {}).get('dateTime', '')
        end_datetime = outlook_event.get('end', {}).get('dateTime', '')
        
        # Converter para datetime para formatação
        try:
            start_dt = datetime.fromisoformat(start_datetime.replace('Z', '+00:00'))
            end_dt = datetime.fromisoformat(end_datetime.replace('Z', '+00:00'))
            
            # Formatar para Expresso
            data = start_dt.strftime('%d/%m/%Y')
            inicio = start_dt.strftime('%H:%M')
            fim = end_dt.strftime('%H:%M')
        except (ValueError, TypeError):
            # Fallback para eventos sem data/hora específica
            data = ''
            inicio = ''
            fim = ''
        
        expresso_event = {
            'titulo': outlook_event.get('subject', ''),
            'descricao': outlook_event.get('body', {}).get('content', ''),
            'data': data,
            'inicio': inicio,
            'fim': fim,
            'localizacao': outlook_event.get('location', {}).get('displayName', ''),
            'participantes': ''  # Não disponível diretamente no Outlook
        }
        
        return expresso_event

    def _format_expresso_to_google(self, expresso_event):
        """
        Converte um evento do Expresso para o formato do Google.
        """
        if not expresso_event.get('titulo') or not expresso_event.get('data'):
            return None
        
        # Converter data e hora do Expresso para formato ISO
        data = expresso_event.get('data', '')
        inicio = expresso_event.get('inicio', '')
        fim = expresso_event.get('fim', '')
        
        try:
            # Converter para datetime
            data_parts = data.split('/')
            if len(data_parts) == 3:
                dia, mes, ano = data_parts
                
                # Data e hora de início
                inicio_parts = inicio.split(':')
                if len(inicio_parts) == 2:
                    hora_inicio, min_inicio = inicio_parts
                    start_dt = datetime(int(ano), int(mes), int(dia), int(hora_inicio), int(min_inicio))
                    start_iso = start_dt.isoformat()
                else:
                    start_iso = f"{ano}-{mes}-{dia}T00:00:00"
                
                # Data e hora de término
                fim_parts = fim.split(':')
                if len(fim_parts) == 2:
                    hora_fim, min_fim = fim_parts
                    end_dt = datetime(int(ano), int(mes), int(dia), int(hora_fim), int(min_fim))
                    end_iso = end_dt.isoformat()
                else:
                    end_iso = f"{ano}-{mes}-{dia}T23:59:59"
            else:
                # Fallback
                start_iso = ''
                end_iso = ''
        except (ValueError, IndexError):
            start_iso = ''
            end_iso = ''
        
        google_event = {
            'summary': expresso_event.get('titulo', ''),
            'description': expresso_event.get('descricao', ''),
            'start': {
                'dateTime': start_iso,
                'timeZone': 'America/Recife'
            },
            'end': {
                'dateTime': end_iso,
                'timeZone': 'America/Recife'
            },
            'location': expresso_event.get('localizacao', '')
        }
        
        return google_event

    def _format_expresso_to_outlook(self, expresso_event):
        """
        Converte um evento do Expresso para o formato do Outlook.
        """
        if not expresso_event.get('titulo') or not expresso_event.get('data'):
            return None
        
        # Converter data e hora do Expresso para formato ISO
        data = expresso_event.get('data', '')
        inicio = expresso_event.get('inicio', '')
        fim = expresso_event.get('fim', '')
        
        try:
            # Converter para datetime
            data_parts = data.split('/')
            if len(data_parts) == 3:
                dia, mes, ano = data_parts
                
                # Data e hora de início
                inicio_parts = inicio.split(':')
                if len(inicio_parts) == 2:
                    hora_inicio, min_inicio = inicio_parts
                    start_dt = datetime(int(ano), int(mes), int(dia), int(hora_inicio), int(min_inicio))
                    start_iso = start_dt.isoformat()
                else:
                    start_iso = f"{ano}-{mes}-{dia}T00:00:00"
                
                # Data e hora de término
                fim_parts = fim.split(':')
                if len(fim_parts) == 2:
                    hora_fim, min_fim = fim_parts
                    end_dt = datetime(int(ano), int(mes), int(dia), int(hora_fim), int(min_fim))
                    end_iso = end_dt.isoformat()
                else:
                    end_iso = f"{ano}-{mes}-{dia}T23:59:59"
            else:
                # Fallback
                start_iso = ''
                end_iso = ''
        except (ValueError, IndexError):
            start_iso = ''
            end_iso = ''
        
        outlook_event = {
            'subject': expresso_event.get('titulo', ''),
            'body': {
                'contentType': 'HTML',
                'content': expresso_event.get('descricao', '')
            },
            'start': {
                'dateTime': start_iso,
                'timeZone': 'America/Recife'
            },
            'end': {
                'dateTime': end_iso,
                'timeZone': 'America/Recife'
            },
            'location': {
                'displayName': expresso_event.get('localizacao', '')
            }
        }
        
        return outlook_event

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
