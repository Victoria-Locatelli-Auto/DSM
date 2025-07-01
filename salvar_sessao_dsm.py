from playwright.sync_api import sync_playwright, TimeoutError

def salvar_sessao():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # navegador visível para login manual
        context = browser.new_context()
        page = context.new_page()

        # Acessa a página de login
        page.goto("https://platoscania-prod.ptcmanaged.com/WebUI/IODealerAnalysis-list.mvc?")

        print("⚠️ Faça login manualmente no navegador. Depois, volte aqui e pressione ENTER para continuar.")
        input()

        # Salva estado da sessão autenticada
        context.storage_state(path="estado_autenticado.json")
        print("✅ Estado de autenticação salvo em 'estado_autenticado.json'.")
        browser.close()

def verificar_autenticacao():
    """ Verifica se o sistema está deslogado e tenta logar novamente """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)  # Executa em segundo plano
        context = browser.new_context(storage_state="estado_autenticado.json")
        page = context.new_page()

        # Acessa a página principal do sistema
        page.goto("https://lpcnet.slapl.prod.aws.scania.com")

        if "login" in page.url:  # Verifica se o sistema redirecionou para a página de login
            print("⚠️ Usuário deslogado! Tentando login automático...")
            page.wait_for_timeout(5000)  # Espera 5 segundos
            page.click("text=Login")  # Clica no botão de login

            # Salva novamente o estado da sessão após login
            context.storage_state(path="estado_autenticado.json")
            print("✅ Reautenticação bem-sucedida!")
        
        browser.close()

if __name__ == "__main__":
    salvar_sessao()  # Primeiro, salva a sessão autenticada
    verificar_autenticacao()  # Em seguida, verifica se há necessidade de login automático