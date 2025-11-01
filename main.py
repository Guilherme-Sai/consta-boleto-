
import time
import re
import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
 

URL_BOLETO = "https://homebanking.srmasset.com/boleto"
ARQUIVO_CNPJS = r'Gui  - boletos/CNPJconsulta.txt'
ARQUIVO_SAIDA = 'relatorio_final_boletos.xlsx'
ARQUIVO_TXT = 'relatorio_final_boletos.txt'
HEADLESS = False
WAIT_AFTER_ACTION = 5
TEMPO_ESPERA_REINICIO = 15
EMAIL_DESTINO = ""  
EMAIL_ORIGEM = "" 
SENHA_EMAIL = "" 
 
def setup_driver(headless: bool = False):
    """Inicia o Chrome."""
    print("[*] Iniciando navegador...")
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(45)
    return driver
 
def consultar_um_cnpj(driver, wait, cnpj):
    """Consulta um único CNPJ e retorna os dados."""
    driver.get(URL_BOLETO)
    input_el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='cpfCnpj']")))
    input_el.clear()
    input_el.send_keys(cnpj)
    input_el.send_keys(Keys.ENTER)
    time.sleep(WAIT_AFTER_ACTION)
 
    try:
        div_duplicatas = driver.find_element(By.CSS_SELECTOR, "div.duplicatas")
        texto_div = div_duplicatas.text
        m = re.search(r'(\d+)', texto_div)
        qtd_boletos = int(m.group(1)) if m else 0
    except:
        qtd_boletos = 0
 
    vencimento = "Não encontrado"
    if qtd_boletos > 0:
        try:
            vencimento_el = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "mat-row:first-of-type mat-cell.mat-column-vencimento"))
            )
            vencimento = vencimento_el.text.strip()
        except:
            vencimento = "Erro ao ler vencimento"
    print(f"  -> {cnpj} | Boletos: {qtd_boletos} | Vencimento: {vencimento}")
    return {'CNPJ_Consultado': cnpj, 'Quantidade de Boletos': qtd_boletos, 'Vencimento Mais Próximo': vencimento}
 
def enviar_email_relatorio(arquivo_txt, email_destino, email_origem, senha):
    """Envia o relatório final em texto por e-mail."""
    print("[*] Enviando relatório final por e-mail...")
    with open(arquivo_txt, 'r', encoding='utf-8') as f:
        conteudo = f.read()
 
    msg = MIMEText(conteudo, 'plain', 'utf-8')
    msg['Subject'] = 'Relatório Final de Boletos'
    msg['From'] = email_origem
    msg['To'] = email_destino
 
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_origem, senha)
        smtp.send_message(msg)
 
    print("[✔] E-mail enviado com sucesso!")
 
def main():
    """Loop principal com tolerância a falhas."""
    try:
        with open(ARQUIVO_CNPJS, 'r', encoding='utf-8') as f:
            cnpjs_total = {re.sub(r'\D', '', l).zfill(14) for l in f if l.strip()}
    except FileNotFoundError:
        print(f"[ERRO] Arquivo não encontrado: {ARQUIVO_CNPJS}")
        return
 
    if os.path.exists(ARQUIVO_SAIDA):
        df = pd.read_excel(ARQUIVO_SAIDA, dtype=str)
    else:
        df = pd.DataFrame(columns=['CNPJ_Consultado', 'Quantidade de Boletos', 'Vencimento Mais Próximo'])
 
    cnpjs_concluidos = set(df['CNPJ_Consultado'])
    cnpjs_pendentes = list(cnpjs_total - cnpjs_concluidos)
    print(f"[*] {len(cnpjs_pendentes)} CNPJs restantes para consulta.")
 
    driver = None
    wait = None
 
    while cnpjs_pendentes:
        try:
            if not driver:
                driver = setup_driver(headless=HEADLESS)
                wait = WebDriverWait(driver, 20)
 
            for cnpj in list(cnpjs_pendentes):
                try:
                    resultado = consultar_um_cnpj(driver, wait, cnpj)
                except Exception as e:
                    resultado = {'CNPJ_Consultado': cnpj, 'Quantidade de Boletos': 'Erro', 'Vencimento Mais Próximo': str(e)}
 
                df = pd.concat([df, pd.DataFrame([resultado])], ignore_index=True)
                df.to_excel(ARQUIVO_SAIDA, index=False)
                print(f"[✓] Progresso salvo: {len(df)} registros.")
                cnpjs_pendentes.remove(cnpj)
 
        except WebDriverException as e:
            print(f"[ERRO GRAVE] Navegador falhou: {e}. Reiniciando em {TEMPO_ESPERA_REINICIO}s...")
            time.sleep(TEMPO_ESPERA_REINICIO)
            if driver:
                driver.quit()
            driver = None
            continue
        except Exception as e:
            print(f"[ERRO] {e}")
            time.sleep(5)
            continue
        finally:
            if driver and not cnpjs_pendentes:
                driver.quit()
 
    
    df.to_excel(ARQUIVO_SAIDA, index=False)
    df.to_csv(ARQUIVO_TXT, sep='\t', index=False, encoding='utf-8')
    print(f"[✔] Relatório final salvo em '{ARQUIVO_TXT}' e '{ARQUIVO_SAIDA}'.")
 
   
    try:
        enviar_email_relatorio(ARQUIVO_TXT, EMAIL_DESTINO, EMAIL_ORIGEM, SENHA_EMAIL)
    except Exception as e:
        print(f"[AVISO] Não foi possível enviar o e-mail automaticamente: {e}")
 
if __name__ == "__main__":
    main()
 
 
