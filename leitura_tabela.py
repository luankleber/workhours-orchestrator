import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def ler_tabela_registros(browser):
    wait = WebDriverWait(browser, 10)

    # Espera a tabela carregar completamente
    tabela = wait.until(
        EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'table')]"))
    )
    linhas = tabela.find_elements(By.XPATH, ".//tr")

    registros = {}

    for linha in linhas[1:]:  # pula o cabeçalho
        colunas = linha.find_elements(By.TAG_NAME, "td")
        if len(colunas) < 3:
            continue

        data_texto = colunas[0].text.strip()
        dia_semana = colunas[1].text.strip()
        turno = colunas[2].text.strip()

        # Inicializa horários
        horarios = []
        for i in range(3, 11):  # colunas de In/Out 1 a 4
            spans = colunas[i].find_elements(By.TAG_NAME, "span")
            hora = None
            for span in spans:
                text = span.text.strip()
                if text and ":" in text:
                    hora = text
                    break
            horarios.append(hora)

        # Mapear para in/out
        registro_horas = {
            "in_1": horarios[0],
            "out_1": horarios[1],
            "in_2": horarios[2],
            "out_2": horarios[3],
            "in_3": horarios[4],
            "out_3": horarios[5],
            "in_4": horarios[6],
            "out_4": horarios[7]
        }

        # Detecta status avançado
        status_coluna = colunas[3]
        status_texto = status_coluna.text.strip()
        status_hover = status_coluna.get_attribute("data-original-title") or ""
        feriado = status_texto.lower() in ["feriado", "holiday"]
        justificado = status_texto.lower() == "justif."
        viagem = status_texto.lower() == "viagem"

        if not any(horarios):
            status = "vazio"
        elif "TAC" in linha.text.upper():
            status = "TAC"
        else:
            status = "ok"

        registros[data_texto] = {
            "dia_semana": dia_semana,
            "turno": turno,
            **registro_horas,
            "status": status,
            "feriado": feriado,
            "justificado": justificado,
            "viagem": viagem,
            "descricao_status": status_hover
        }

    # Salva como JSON
    with open("registros_mensais.json", "w", encoding="utf-8") as f:
        json.dump(registros, f, indent=2, ensure_ascii=False)

    print(f"✅ {len(registros)} registros extraídos e salvos em registros_mensais.json")
    return registros
