---

# Automação de Relatórios Plato Scania com Envio por Email

Este projeto automatiza a geração diária de relatórios da plataforma **Plato Scania**, baixando os dados em CSV, convertendo para XLSX e enviando os arquivos por email para destinatários pré-configurados. Além disso, inclui scripts para gerenciar a autenticação na plataforma usando Playwright, salvando e verificando sessões autenticadas. Também previne que o computador entre em modo de suspensão durante a execução da rotina.

---

## Funcionalidades Principais

### 1. Geração e Envio de Relatórios

* Acessa a plataforma Plato Scania via browser automatizado (Playwright).
* Aguarda a tabela de dados carregar e realiza download do relatório em CSV.
* Converte automaticamente o CSV para XLSX, organizando os dados por colunas específicas.
* Envia os relatórios gerados por email para uma lista de destinatários.
* Mantém o computador ativo para evitar suspensão durante a execução.
* Agenda a rotina para rodar automaticamente todos os dias em horário configurável.

### 2. Gerenciamento de Sessão de Login com Playwright

* Permite realizar login manual no navegador para salvar o estado da sessão autenticada (`estado_autenticado.json`).
* Verifica automaticamente se a sessão expirou e tenta relogar para manter a sessão ativa.
* Facilita a automação segura sem necessidade de inserir login/senha toda vez.

---

## Pré-requisitos

* Python 3.8 ou superior
* Bibliotecas Python:

  * `playwright`
  * `pandas`
  * `openpyxl`
* Navegador Chromium instalado (Playwright instala automaticamente se necessário)
* Conta de email com senha de aplicativo para envio SMTP
* Arquivo `estado_autenticado.json` contendo a sessão autenticada no Plato Scania (gerado pelo script de sessão)

---

## Instalação

1. Clone este repositório ou copie os scripts para seu ambiente.

2. Instale as dependências:

```bash
pip install pandas openpyxl playwright
playwright install
```

3. Gere uma senha de aplicativo para o email remetente (no Outlook/Gmail) e configure no script principal:

```python
SENHA_EMAIL = "sua_senha_de_aplicativo_aqui"
```

4. Ajuste os emails remetente e destinatários conforme necessário.

---

## Uso

### Passo 1: Salvar Sessão Autenticada (Login Manual)

Execute o script de sessão para abrir o navegador e fazer login manual no Plato Scania:

```bash
python salvar_sessao.py
```

* O navegador abrirá para que você realize o login.
* Após login, volte ao terminal e pressione ENTER.
* O estado da sessão será salvo em `estado_autenticado.json`.

### Passo 2: Verificar Autenticação

O script de verificação pode ser executado para checar se a sessão está ativa e relogar se necessário:

```bash
python verificar_autenticacao.py
```

### Passo 3: Executar Rotina Automática de Relatórios

Execute o script principal que agenda a rotina diária de geração e envio dos relatórios:

```bash
python gerar_relatorio_dsm.py
```

O script ficará rodando e executará a rotina no horário programado (exemplo: 08:32).

---

## Configurações Importantes

* **EMAIL\_REMETENTE**: Email que envia os relatórios.
* **SENHA\_EMAIL**: Senha de aplicativo para autenticação SMTP.
* **EMAIL\_DESTINATARIO**: Lista de emails que receberão os relatórios.
* **PASTA\_RELATORIOS**: Diretório para salvar os arquivos CSV e XLSX.
* **SMTP\_SERVIDOR** e **SMTP\_PORTA**: Configurações do servidor SMTP.
* **agendar\_rotina("HH\:MM")**: Horário para execução diária da rotina.

---

## Como funciona a prevenção da suspensão do PC?

O script usa chamadas do Windows API via `ctypes` para manter o sistema ativo durante a execução, evitando que o computador entre em modo de suspensão.

---

## Segurança

* Utilize senhas de aplicativo específicas para evitar expor sua senha principal.
* Proteja o arquivo `estado_autenticado.json`, pois ele contém dados de sessão autenticada.

---

## Logs e Tratamento de Erros

* Mensagens de status e erro são exibidas no console durante a execução.
* Possíveis falhas na geração do relatório ou envio de email são notificadas.

---

## Possíveis Melhorias Futuras

* Configuração via arquivo `.env` ou arquivo de configuração externa.
* Registro de logs em arquivo para histórico.
* Notificações em caso de falhas via outros canais (Slack, Telegram etc).
* Interface gráfica simples para controlar a automação.

---
