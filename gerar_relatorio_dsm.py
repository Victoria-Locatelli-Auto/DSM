import smtplib
import os
import pandas as pd
import time
import ctypes
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from playwright.sync_api import sync_playwright

# ==================== 🔌 EVITAR SUSPENSÃO ====================
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_AWAYMODE_REQUIRED = 0x00000040

def manter_pc_ativo():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
    )

def permitir_suspensao():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# ==================== CONFIGURAÇÕES ==========================

EMAIL_REMETENTE = "@@@@@"
SENHA_EMAIL = "@@@@@"  # ⚠ Substitua aqui pela senha de aplicativo gerada no Outlook!
EMAIL_DESTINATARIO = ["@@@@@"]
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PORTA = 587

PASTA_RELATORIOS = "relatorios"
os.makedirs(PASTA_RELATORIOS, exist_ok=True)

# ==================== FUNÇÕES ==========================

def obter_data_formatada():
    return datetime.now().strftime("%d-%m-%Y")

def gerar_relatorio_plato(context):
    page = context.new_page()

    try:
        print("➡️ Acessando Plato Scania...")
        page.goto("https://platoscania-prod.ptcmanaged.com/WebUI/IODealerAnalysis-list.mvc")

        page.wait_for_load_state('networkidle')
        page.wait_for_timeout(5000)

        print("✅ Página carregada!")

        print("⏳ Aguardando a tabela dinâmica carregar...")
        page.wait_for_selector('table.full.grid-inline-edit', timeout=120000)
        print("✅ Tabela detectada.")

        print("⏳ Aguardando botão de exportação...")
        export_button = page.locator('li#layout-control-export').nth(1)
        export_button.wait_for(state="visible", timeout=120000)
        print("✅ Botão de exportação visível.")

        print("🚀 Iniciando download do CSV...")

        with page.expect_download() as download_info:
            export_button.click()

        download = download_info.value

        nome_base = f"DSM_{obter_data_formatada()}"
        caminho_csv = os.path.join(PASTA_RELATORIOS, f"{nome_base}.csv")
        caminho_xlsx = os.path.join(PASTA_RELATORIOS, f"{nome_base}.xlsx")

        download.save_as(caminho_csv)
        print(f"✅ CSV salvo como {caminho_csv}")

        # 📄 Converter CSV em XLSX
        print("🔄 Convertendo CSV para XLSX...")

        try:
            df = pd.read_csv(caminho_csv, skiprows=3, sep=',', encoding='utf-8')

            if 'Part' in df.columns and 'Location' in df.columns:
                df = df.sort_values(by=['Part', 'Location'], ascending=[True, True])
                print("✅ Dados ordenados por Part e Location.")
            else:
                print("⚠️ Colunas 'Part' e/ou 'Location' não encontradas para ordenação.")

            df.to_excel(caminho_xlsx, index=False, engine='openpyxl')
            print(f"✅ XLSX salvo como {caminho_xlsx}")

        except Exception as e:
            print(f"❌ Erro na conversão CSV -> XLSX: {e}")

        return [caminho_csv, caminho_xlsx]

    except Exception as e:
        print(f"❌ Erro ao gerar relatório: {e}")
        return []


def enviar_email(anexos):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = ", ".join(EMAIL_DESTINATARIO)
    msg['Subject'] = f"Relatórios DSM - {obter_data_formatada()}"

    corpo = MIMEText("Segue em anexo os relatórios gerados automaticamente em CSV e XLSX.", 'plain')
    msg.attach(corpo)

    for arquivo in anexos:
        with open(arquivo, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(arquivo))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo)}"'
            msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_REMETENTE, SENHA_EMAIL)
            server.send_message(msg)

        print("📧 Email enviado com sucesso!")

    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ Erro de autenticação SMTP: {e}")
        print("⚠️ Verifique se sua senha de aplicativo está correta e se sua conta permite envio SMTP.")
    except Exception as e:
        print(f"❌ Erro ao enviar email: {e}")

def executar_rotina():
    manter_pc_ativo()

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(storage_state="estado_autenticado.json")

            arquivos_gerados = gerar_relatorio_plato(context)

            browser.close()
            
            
            if arquivos_gerados:
                enviar_email(arquivos_gerados)
            else:
                print("⚠️ Nenhum relatório foi gerado para enviar.")
            

    finally:
        permitir_suspensao()

def agendar_rotina(hora_agendada="08:32"):
    print(f"⏰ Aguardando para rodar todos os dias às {hora_agendada}...")

    while True:
        agora = datetime.now()
        hora_execucao = datetime.strptime(hora_agendada, "%H:%M").replace(
            year=agora.year, month=agora.month, day=agora.day
        )

        if agora >= hora_execucao:
            hora_execucao += timedelta(days=1)

        tempo_espera = (hora_execucao - agora).total_seconds()

        print(f"🕒 Próxima execução em {hora_execucao.strftime('%d/%m/%Y %H:%M')}")

        time.sleep(tempo_espera)

        print("🚀 Iniciando execução da rotina!")
        executar_rotina()

if __name__ == "__main__":
    agendar_rotina("08:32")
