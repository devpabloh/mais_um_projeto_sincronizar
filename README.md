# Sincronizador de Calendários

Este projeto permite a sincronização de eventos entre Google Calendar, Outlook e Expresso.

## Pré-requisitos

- Python 3.8 ou superior
- Conta no Google Cloud Platform
- Conta Microsoft 365 (para Outlook)
- Conta no Expresso (se necessário)

## Configuração Inicial

1. **Clone o repositório**
   ```bash
   git clone [URL_DO_REPOSITÓRIO]
   cd mais_um_projeto_sincronizar
   ```

2. **Crie um ambiente virtual (recomendado)**
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # No Windows
   # ou
   source .venv/bin/activate  # No Linux/Mac
   ```

3. **Instale as dependências**
   ```bash
   pip install -r requirements.txt
   ```

## Configuração das Credenciais

### Google Calendar

1. Acesse o [Google Cloud Console](https://console.cloud.google.com/)
2. Crie um novo projeto ou selecione um existente
3. Ative a API do Google Calendar
4. Vá para "Credenciais" > "Criar Credenciais" > "ID do cliente OAuth"
5. Selecione "Aplicativo da área de trabalho"
6. Faça o download do arquivo JSON e renomeie para `credentials.json` na raiz do projeto

### Outlook Calendar

1. Acesse o [Portal do Azure](https://portal.azure.com/)
2. Vá para "Azure Active Directory" > "Registros de aplicativos" > "Novo registro"
3. Configure o aplicativo:
   - Nome: SeuAppSincronizador
   - Tipos de conta contas em qualquer diretório organizacional
   - URI de redirecionamento: `http://localhost:50141`
4. Anote os seguintes valores:
   - ID do aplicativo (client_id)
   - ID do diretório (tenant_id)
5. Vá para "Certificados e segredos" e crie um novo segredo do cliente
6. Anote o valor do segredo (client_secret)

## Configuração do Código

1. Abra o arquivo `main.py`
2. Localize a seção de configuração do Outlook (por volta da linha 20) e atualize com suas credenciais:
   ```python
   outlook_client_id = "SEU_CLIENT_ID_AQUI"
   outlook_client_secret = "SEU_CLIENT_SECRET_AQUI"
   outlook_tenant_id = "SEU_TENANT_ID_AQUI"
   outlook_calendar_id = "SEU_CALENDAR_ID_AQUI"
   ```

## Primeira Execução

1. Execute o script principal:
   ```bash
   python main.py
   ```
2. Na primeira execução, você será redirecionado para fazer login nas contas do Google e Outlook
3. Acesse as URLs fornecidas no terminal para autorizar o acesso
4. Após a autorização, os tokens serão salvos automaticamente nos arquivos `token.json`

## Uso

A aplicação irá sincronizar automaticamente os eventos entre os calendários configurados. Por padrão, ela:

1. Busca eventos futuros do Outlook
2. Busca eventos futuros do Google Calendar
3. Sincroniza os eventos entre as plataformas
4. Remove duplicatas e conflitos

## Personalização

Você pode ajustar o comportamento da sincronização editando o arquivo `main.py`:

- `lookahead_days`: Número de dias futuros para sincronizar
- `event_buffer_minutes`: Tempo de buffer entre eventos
- `max_retries`: Número de tentativas em caso de falha

## Solução de Problemas

- **Erro de autenticação**: Verifique se as credenciais estão corretas e se as APIs estão ativadas
- **Problemas de permissão**: Certifique-se de que concedeu todas as permissões necessárias
- **Erros de conexão**: Verifique sua conexão com a internet

## Segurança

- Nunca compartilhe seus arquivos `credentials.json` e `token.json`
- Adicione esses arquivos ao `.gitignore`
- Mantenha suas credenciais em segredo

## Suporte

Em caso de problemas, abra uma issue no repositório do projeto.
