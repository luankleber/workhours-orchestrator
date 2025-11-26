from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime
import json
import time
import random

def fechar_modal_se_existir(browser, wait):
    try:
        modal = browser.find_element(By.ID, "modal_add_register")
        if modal.is_displayed():
            btn_fechar = modal.find_element(By.CSS_SELECTOR, "button[data-dismiss='modal']")
            btn_fechar.click()
            wait.until(EC.invisibility_of_element(modal))
            print("⚠️ Modal anterior fechado")
    except NoSuchElementException:
        pass

def preencher_modal(browser, plano_json_path="plano_para_preenchimento.json"):
    wait = WebDriverWait(browser, 10)

    with open(plano_json_path, "r", encoding="utf-8") as f:
        plano = json.load(f)

    dias = sorted(plano.keys(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"))

    for data_str in dias:
        info = plano[data_str]
        if info.get("status") != "vazio":
            continue  # só preenche dias vazios

        horarios = [info.get("in_1"), info.get("out_1"), info.get("in_2"), info.get("out_2")]
        horarios = [h for h in horarios if h]  # remove None ou vazio

        for horario in horarios:
            try:
                fechar_modal_se_existir(browser, wait)

                # localiza a linha do dia
                linha = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//tr[td/span[contains(text(), '{data_str}')]]")
                    )
                )

                # clica no botão "+" para abrir modal
                botao_add = wait.until(
                    EC.element_to_be_clickable(
                        linha.find_element(By.XPATH, ".//div[contains(@onclick, 'addRegister')]")
                    )
                )
                browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_add)
                botao_add.click()

                # espera modal abrir
                modal = wait.until(
                    EC.visibility_of_element_located((By.ID, "modal_add_register"))
                )
                print(f"🔹 Modal aberto para {data_str}, horário {horario}")

                # preenche horário
                input_horario = modal.find_element(By.ID, "time_add_register")
                browser.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                    input_horario, horario
                )
                print(f"   ⏰ Horário preenchido: {horario}")

                # preenche observação
                input_obs = modal.find_element(By.ID, "obs_add_register")
                input_obs.clear()
                input_obs.send_keys("Trabalho")
                print(f"   📝 Observação preenchida: Trabalho")

                # --- Clica em fechar (por enquanto, não salva) ---
                #btn_fechar = modal.find_element(By.CSS_SELECTOR, "button[data-dismiss='modal']")
                #btn_fechar.click()
                #wait.until(EC.invisibility_of_element(modal))
                #print(f"⚠️ Modal fechado para {data_str}, horário {horario}")

                # --- Salva os horarios
                btn_save = modal.find_element(By.ID, "modal_add_register_save")
                btn_save.click()
                wait.until(EC.invisibility_of_element(modal))
                print(f"⚠️ Horario Salvo para {data_str}, horário {horario}")

                # --- Delay aleatório para simular comportamento humano ---
                time.sleep(random.uniform(1, 3))

            except TimeoutException:
                print(f"❌ Elemento não encontrado para {data_str} / horário {horario}")
            except Exception as e:
                print(f"❌ Erro ao preencher modal para {data_str} / horário {horario}: {e}")

    print("🎉 Todos os registros faltantes preparados (modais preenchidos e fechados)!")
