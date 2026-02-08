import time
import smtplib
import datetime
import os
import re
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES DE AMBIENTE (Segurança) ---
EMAIL_REMETENTE = os.environ.get("EMAIL_USER")
SENHA_APP = os.environ.get("EMAIL_PASS")
EMAIL_DESTINATARIO = os.environ.get("EMAIL_DEST")

URL_BI = "https://app.powerbi.com/view?r=eyJrIjoiOWNmNWY5YWQtNGUzYi00OTAxLTgwMDUtYmExN2Q4YTA0ZDNmIiwidCI6Ijg5MzJiNTAxLTRkMTQtNGIyOC04ZGUxLTg4YjgzYThiN2MwZCJ9&pageName=a45d0354e465654433c3"
# URL DA PLANILHA (Pega a aba 1 usando gid=0)
URL_PLANILHA = "https://docs.google.com/spreadsheets/d/1tjQJn9IFeALBVELxFOqleBd5lky4Wm3KGanX0LCFOdI/edit?gid=0#gid=0"

MESES_REV = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}

def forcar_clique(driver, elemento):
    driver.execute_script("arguments[0].click();", elemento)

def ler_meta_planilha_h59():
    print("\n--- 📊 INICIANDO LEITURA DA PLANILHA (ALVO: CÉLULA B1) ---")
    try:
        sheet_id = URL_PLANILHA.split('/d/')[1].split('/')[0]
        url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0"
        
        df = pd.read_csv(url_csv, header=None, dtype=str, skip_blank_lines=False)
        
        linha_alvo = 0 
        coluna_alvo = 1 
        
        if len(df) > linha_alvo:
            valor_bruto = str(df.iloc[linha_alvo, coluna_alvo])
            if valor_bruto and valor_bruto.lower() != 'nan':
                valor_limpo = valor_bruto.replace('R$', '').replace(' ', '').replace('.', '')
                valor_limpo = re.sub(r'[^\d,]', '', valor_limpo)
                print(f"✅ VALOR CAPTURADO (B1): {valor_limpo}")
                return valor_limpo
        return None
    except Exception as e:
        print(f"Erro planilha: {e}")
        return None

def get_datas_filtro():
    hoje = datetime.date.today()
    data_alvo_slicer = hoje.replace(day=1)
    mes_anterior = data_alvo_slicer - datetime.timedelta(days=1)
    meses_map = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho", 
                 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
    return str(mes_anterior.year), meses_map[mes_anterior.month], data_alvo_slicer

def encontrar_elemento_em_frames(driver, by, locator):
    elementos = driver.find_elements(by, locator)
    if elementos: return elementos[0]
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        try:
            driver.switch_to.frame(frame)
            res = encontrar_elemento_em_frames(driver, by, locator)
            if res: return res
            driver.switch_to.parent_frame()
        except: driver.switch_to.default_content()
    return None

def aplicar_filtro_sherlock(driver, nome_filtro, valor_desejado):
    print(f"--- Filtro '{nome_filtro}' -> '{valor_desejado}' ---")
    xpath_slicer = f"//*[@aria-label='{nome_filtro}'] | //*[@title='{nome_filtro}'] | //div[contains(@class, 'header') and .//text()='{nome_filtro}']"
    slicer = encontrar_elemento_em_frames(driver, By.XPATH, xpath_slicer)
    if not slicer: return
    try:
        container = slicer.find_element(By.XPATH, "./ancestor::div[contains(@class, 'visualContainer')][1]")
        try:
            seta = container.find_element(By.CSS_SELECTOR, "i.chevron-down, .slicer-dropdown-menu")
            forcar_clique(driver, seta)
        except: forcar_clique(driver, container)
        time.sleep(3)
        xpath_val = f"//span[@title='{valor_desejado}'] | //span[text()='{valor_desejado}']"
        try:
            item = driver.find_element(By.XPATH, xpath_val)
            driver.execute_script("arguments[0].scrollIntoView(true);", item)
            time.sleep(1)
            forcar_clique(driver, item)
            print("Selecionado.")
        except: print("Valor não encontrado na lista.")
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(2)
    except: pass

def ajustar_data_calendario(driver, data_alvo):
    print(f"--- Ajustando Calendário para: {data_alvo.strftime('%d/%m/%Y')} ---")
    xpath_visual = "//*[contains(@class,'visualContainer')][.//text()='Data - Período' or @title='Data - Período']"
    visual = encontrar_elemento_em_frames(driver, By.XPATH, xpath_visual)
    if not visual: return
    try:
        inputs = visual.find_elements(By.TAG_NAME, "input")
        input_fim = [i for i in inputs if i.is_displayed()][1]
        driver.execute_script("arguments[0].click();", input_fim)
        time.sleep(3.0)
        xpath_cal = "//div[contains(@class, 'calendar') and contains(@class, 'themeableElement')]"
        cal_container = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, xpath_cal)))
        target_mes, target_ano = data_alvo.month, data_alvo.year
        for _ in range(12):
            texto = cal_container.text.lower()
            mes_cal = next((num for nome, num in MESES_REV.items() if nome in texto), None)
            ano_match = re.search(r'\d{4}', texto)
            ano_cal = int(ano_match.group(0)) if ano_match else None
            if not mes_cal or not ano_cal: break
            val_atual, val_target = ano_cal * 100 + mes_cal, target_ano * 100 + target_mes
            if val_atual == val_target: break
            btn_xpath = ".//button[contains(@aria-label, 'Próximo')]" if val_atual < val_target else ".//button[contains(@aria-label, 'Anterior')]"
            ActionChains(driver).move_to_element(cal_container.find_element(By.XPATH, btn_xpath)).click().perform()
            time.sleep(1.0)
        xpath_dia = ".//*[contains(@class, 'date-cell') and normalize-space(text())='1']"
        dias = cal_container.find_elements(By.XPATH, xpath_dia)
        dia_alvo_elem = next((d for d in dias if d.is_displayed()), None)
        if dia_alvo_elem:
            ActionChains(driver).move_to_element(dia_alvo_elem).click().perform()
            time.sleep(2)
            print("SUCESSO: Dia 1 selecionado.")
            time.sleep(3)
    except: pass

def ajustar_meta_loja(driver, valor):
    if not valor: return
    print(f"--- Inserindo Meta: {valor} ---")
    xpath_vis = "//*[contains(@class,'visualContainer')][.//text()='MetaLoja' or @title='MetaLoja']"
    visual = encontrar_elemento_em_frames(driver, By.XPATH, xpath_vis)
    if not visual: return
    try:
        inputs = visual.find_elements(By.TAG_NAME, "input")
        for campo in [i for i in inputs if i.is_displayed()]:
            ActionChains(driver).click(campo).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.DELETE).perform()
            time.sleep(0.5)
            campo.send_keys(valor)
            time.sleep(0.5)
            campo.send_keys(Keys.TAB)
            time.sleep(1.0)
    except: pass

def extrair_tabela(driver, tabela_element):
    celulas = tabela_element.find_elements(By.CSS_SELECTOR, ".pivotTableCellWrap, .ui-grid-cell-contents")
    lista = [c.text.strip() for c in celulas if c.text.strip() != ""]
    html = """<table style="border-collapse: collapse; width: 600px; font-family: Arial; border: 1px solid #ddd;">
    <tr style="background-color: #0f4c3a; color: white;"><th style="padding: 10px;">Vendedor</th><th style="padding: 10px;">Comissão</th><th style="padding: 10px;">Prêmiação</th></tr>"""
    blacklist = ["Meta", "Bonus", "TKM", "PMA", "Tot.", "SEM VENDEDOR", "Vendedor", "Comissão", "Premiação", "Prêmiação", "IPA", "Gorjeta"]
    i = 0
    while i < len(lista):
        item = lista[i]
        if item == "Total":
            vals = [x for x in lista[i+1:i+6] if "R$" in x or (any(c.isdigit() for c in x) and "," in x)]
            val1, val2 = (vals[0] if len(vals)>=1 else "-", vals[1] if len(vals)>=2 else "-")
            html += f"<tr style='background-color: #e6f2ef; font-weight: bold;'><td style='padding:8px;'>Total</td><td style='padding:8px;'>{val1}</td><td style='padding:8px;'>{val2}</td></tr>"
            break
        if item.startswith("R$") or (len(item)>0 and item[0].isdigit()) or any(b in item for b in blacklist):
            i+=1; continue
        vals = []
        for x in lista[i+1:i+12]:
            if "R$" in x or (any(c.isdigit() for c in x) and "," in x and len(x)<20): vals.append(x)
            if len(vals)==2: break
        val1, val2 = (vals[0] if len(vals)>=1 else "-", vals[1] if len(vals)>=2 else "-")
        html += f"<tr style='border-bottom: 1px solid #eee;'><td style='padding:8px;'>{item}</td><td style='padding:8px;'>{val1}</td><td style='padding:8px;'>{val2}</td></tr>"
        i+=1
    return html + "</table>"

def extrair_tabela_gorjeta(driver, tabela_element):
    if not tabela_element: return ""
    celulas = tabela_element.find_elements(By.CSS_SELECTOR, ".pivotTableCellWrap, .ui-grid-cell-contents")
    lista = [c.text.strip() for c in celulas if c.text.strip() != ""]
    
    html = """<table style="border-collapse: collapse; width: 400px; font-family: Arial; border: 1px solid #ddd;">
    <tr style="background-color: #0f4c3a; color: white;"><th style="padding: 10px;">Vendedor</th><th style="padding: 10px;">Gorjeta</th></tr>"""
    
    blacklist = ["SEM VENDEDOR", "Vendedor", "Gorjeta", "R$ Total", "TKM", "PMA", "IPA", "Meta", "Bonus", "Premiação", "Comissão", "Tot.", "Arom.", "Puro", "Acess."]
    
    i = 0
    while i < len(lista):
        item = lista[i]
        
        if any(bad in item for bad in blacklist) or (item.startswith("R$") or (len(item)>0 and item[0].isdigit())):
            i += 1
            continue
            
        if item == "Total":
            val = lista[i+1] if i+1 < len(lista) else "-"
            html += f"<tr style='background-color: #e6f2ef; font-weight: bold;'><td style='padding:8px;'>Total</td><td style='padding:8px;'>{val}</td></tr>"
            break
            
        vendedor = item
        gorjeta = "-"
        if i+1 < len(lista):
            prox = lista[i+1]
            if "R$" in prox or (any(c.isdigit() for c in prox) and "," in prox):
                gorjeta = prox
            
        html += f"<tr style='border-bottom: 1px solid #eee;'><td style='padding:8px;'>{vendedor}</td><td style='padding:8px;'>{gorjeta}</td></tr>"
        i+=1 
        
    return html + "</table>"

def enviar_email(anexo, mes, ano, html_comissao, html_gorjeta, meta_valor):
    print("--- Preparando envio de e-mail ---")
    if not EMAIL_REMETENTE or not SENHA_APP:
        print("FALHA: Credenciais de e-mail não encontradas nas variáveis de ambiente.")
        return
        
    try:
        msg = MIMEMultipart('related')
        msg['Subject'] = f"[TS Vila Madalena] Comissões e Gorjetas - {mes}/{ano}"
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINATARIO
        texto_meta = f"R$ {meta_valor}" if meta_valor else "Não capturada"
        
        html = f"""<html><body>
        <h2 style='color:#0f4c3a;'>Relatório de Comissionamento</h2>
        <p>Ref: <b>{mes}/{ano}</b></p>
        <p><b>Meta da Loja: {texto_meta}</b></p>
        <br>
        {html_comissao}
        <br>
        <h3 style='color:#0f4c3a;'>Relatório de Gorjetas</h3>
        {html_gorjeta}
        <br>
        <p style="font-family: Arial; font-size: 12px; color: gray;"><i>O print original segue em anexo.</i></p>
        </body></html>"""
        
        msg.attach(MIMEText(html, 'html'))
        with open(anexo, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-Disposition', 'attachment', filename=anexo)
            msg.attach(img)
            
        print(f"Conectando ao servidor SMTP (User: {EMAIL_REMETENTE})...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP)
        server.send_message(msg)
        server.quit()
        print("✅ E-mail enviado com sucesso!")
        
    except Exception as e:
        print(f"❌ ERRO AO ENVIAR E-MAIL: {e}")

def executar_robo():
    valor_meta = ler_meta_planilha_h59()
    ano_dd, mes_dd, data_alvo = get_datas_filtro()
    print(f"Iniciando: {mes_dd}/{ano_dd}. Alvo Slicer: {data_alvo.strftime('%d/%m/%Y')}. Meta: {valor_meta}")

    opts = webdriver.ChromeOptions()
    opts.add_argument("--headless"); opts.add_argument("--no-sandbox"); opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--lang=pt-BR")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    arq = f"relatorio_{mes_dd}_{ano_dd}.png"
    try:
        driver.get(URL_BI); time.sleep(30); driver.switch_to.default_content()
        aplicar_filtro_sherlock(driver, "Ano", ano_dd); driver.switch_to.default_content()
        aplicar_filtro_sherlock(driver, "Mês", mes_dd); driver.switch_to.default_content()
        ajustar_data_calendario(driver, data_alvo)
        if valor_meta: driver.switch_to.default_content(); ajustar_meta_loja(driver, valor_meta)
        
        time.sleep(5); driver.switch_to.default_content()
        
        # 1. Tabela Comissão
        xp_comissao = "//div[contains(@class,'visualContainer')][descendant::*[contains(text(), 'Premiação')]]"
        tab_comissao = encontrar_elemento_em_frames(driver, By.XPATH, xp_comissao)
        html_comissao = extrair_tabela(driver, tab_comissao) if tab_comissao else "<p>Erro tab. comissão</p>"
        
        driver.switch_to.default_content()
        
        # 2. Tabela Gorjeta
        xp_gorjeta = "//div[contains(@class,'visualContainer')][descendant::*[contains(text(), 'Gorjeta')]]"
        tabelas_possiveis = driver.find_elements(By.XPATH, xp_gorjeta)
        
        tab_gorjeta = None
        for t in tabelas_possiveis:
            if "Premiação" not in t.get_attribute("textContent"):
                tab_gorjeta = t
                break
        
        if not tab_gorjeta and tabelas_possiveis: tab_gorjeta = tabelas_possiveis[0]
            
        html_gorjeta = extrair_tabela_gorjeta(driver, tab_gorjeta) if tab_gorjeta else "<p>Erro tab. gorjeta</p>"
        
        if tab_comissao: tab_comissao.screenshot(arq)
        else: driver.save_screenshot(arq)
        
        return arq, mes_dd, ano_dd, html_comissao, html_gorjeta, valor_meta
    finally: driver.quit()

if __name__ == "__main__":
    try:
        a, m, y, h_comissao, h_gorjeta, meta = executar_robo()
        enviar_email(a, m, y, h_comissao, h_gorjeta, meta)
    except Exception as e: print(f"Erro Fatal: {e}")



