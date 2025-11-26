import json
from dotenv import load_dotenv
import os
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from leitura_tabela import ler_tabela_registros
from gerar_plano import gerar_plano
from preencher_registros import preencher_modal

# --- Configurações do Chrome ---
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)

browser = webdriver.Chrome(service=Service(), options=chrome_options)
wait = WebDriverWait(browser, 15)

# --- Abre o site e faz login ---
browser.get("https://www.pmovel.com.br/")
try:
    email_input = wait.until(EC.presence_of_element_located((By.NAME, "email")))
    senha_input = browser.find_element(By.NAME, "password")
    botao_login = browser.find_element(
        By.XPATH,
        "//form//button[@type='submit' or contains(@class, 'btn')]"
    )

    email_input.send_keys(os.getenv("PMOVEL_USER"))
    senha_input.send_keys(os.getenv("PMOVEL_PASS"))
    botao_login.click()

except Exception as e:
    print("❌ Erro ao tentar logar:", e)
    browser.quit()
    raise

# --- Aguarda menu principal e acessa 'Registros' ---
try:
    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//ul//a[contains(translate(., 'REGISTROS', 'registros'), 'registros')]")
        )
    )
    print("✅ Login bem-sucedido e menu carregado!")

    menu_registros = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(translate(., 'REGISTROS', 'registros'), 'registros')]")
        )
    )
    menu_registros.click()

    # Captura número de linhas iniciais
    tabela_inicial = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//table[contains(@class, 'table') or contains(@id, 'registros')]")
        )
    )
    linhas_iniciais = tabela_inicial.find_elements(By.XPATH, ".//tbody/tr")
    qtd_inicial = len(linhas_iniciais)

    # Tenta selecionar 'Mês Atual' no dropdown
    try:
        range_btn = WebDriverWait(browser, 3).until(
            EC.element_to_be_clickable((By.ID, "Areportrange"))
        )
        range_btn.click()

        mes_atual_item = WebDriverWait(browser, 3).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//ul/li[contains(@data-range-key, 'Mês Atual') or contains(@data-range-key, 'This Month')]")
            )
        )
        mes_atual_item.click()

        aplicar_btn = WebDriverWait(browser, 3).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'applyBtn') and (text()='Aplicar' or text()='Apply')]")
            )
        )
        aplicar_btn.click()
        wait.until(lambda b: len(b.find_elements(By.XPATH, "//table[contains(@class, 'table')]/tbody/tr")) > qtd_inicial)
        print("✅ Dropdown 'Mês Atual' selecionado e tabela completa carregada!")

    except:
        pass

except Exception as e:
    print("⚠️ Erro ao acessar 'Meus registros', mas continuando:", e)

# --- Lê os registros da tabela ---
registros = ler_tabela_registros(browser)

# --- FASE 2: Geração automática de plano ---
from gerar_plano import gerar_plano
plano_completo = gerar_plano()

# --- FASE 3: Preparar registros para preenchimento ---
from preencher_registros import preencher_modal
preencher_modal(browser, plano_json_path="plano_para_preenchimento.json")