import win32com.client
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import tempfile, os

# --- Função para pegar o HTML do email ---
def get_latest_email_html(subject_filter):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Caixa de entrada

    messages = inbox.Items
    valid_messages = []

    for m in messages:
        try:
            _ = m.ReceivedTime
            valid_messages.append(m)
        except Exception:
            continue

    if not valid_messages:
        return None

    valid_messages = sorted(valid_messages, key=lambda x: x.ReceivedTime, reverse=True)

    for mail in valid_messages:
        try:
            if subject_filter.lower() in mail.Subject.lower():
                return mail.HTMLBody
        except Exception:
            continue

    return None

# --- Função para extrair a tabela renderizada com Selenium ---
def extract_salesforce_table(html):
    tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    tmpfile.write(html.encode('utf-8'))
    tmpfile.close()

    chrome_opts = Options()
    chrome_opts.add_argument("--headless")
    chrome_opts.add_argument("--disable-gpu")
    chrome_opts.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=chrome_opts)
    driver.get("file:///" + tmpfile.name)

    # O Salesforce geralmente usa classes para tabelas, mas podemos pegar a primeira table visível
    tables = driver.find_elements("tag name", "table")
    data = []

    for table in tables:
        try:
            rows = table.find_elements("tag name", "tr")
            for r in rows:
                cells = r.find_elements("xpath", ".//th|.//td")
                row_data = [c.text.strip() for c in cells if c.text.strip() != ""]
                if row_data:
                    data.append(row_data)
            if data:
                break  # Pega só a primeira tabela válida
        except Exception:
            continue

    driver.quit()
    os.unlink(tmpfile.name)
    return data

# --- Função para gerar Excel só com as colunas que importam ---
def generate_excel(subject="Relatar resultados (Tabela de Horas Trabalhadas)"):
    html = get_latest_email_html(subject)
    if not html:
        print("❌ Nenhum email encontrado.")
        return

    table_data = extract_salesforce_table(html)
    if not table_data:
        print("❌ Nenhuma tabela parseável encontrada.")
        return

    # --- Filtrando apenas as colunas que queremos ---
    headers = table_data[0]
    wanted_columns = [
        "Hora de início↓",
        "Hora de término",
        "Time Entry Type",
        "Service Appointment: Account Name | Site Name"
    ]

    col_indexes = [headers.index(c) for c in wanted_columns if c in headers]
    if not col_indexes:
        print("❌ As colunas desejadas não foram encontradas.")
        return

    filtered_data = [[row[i] for i in col_indexes] for row in table_data[1:]]
    df = pd.DataFrame(filtered_data, columns=wanted_columns)

    filename = f"salesforce_{subject.replace(' ','_')}.xlsx"
    df.to_excel(filename, index=False)
    print(f"✅ Excel gerado: {filename}")

# --- Executa ---
if __name__ == "__main__":
    generate_excel()
