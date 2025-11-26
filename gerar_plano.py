import json
from datetime import datetime, timedelta

def gerar_plano():
    # --- Carrega registros existentes ---
    with open("registros_mensais.json", "r", encoding="utf-8") as f:
        registros = json.load(f)

    # --- Configurações de horários ---
    hora_entrada = "07:30"
    hora_saida = "16:54"
    tempo_almoco = 1  # 1 hora, já considerado pelo sistema
    limite_horas = 10
    hoje = datetime.today().date()

    def str_para_data(s):
        return datetime.strptime(s, "%d/%m/%Y").date()

    def calcular_horas_liquidas(inicio, fim, almoco):
        h_in, m_in = map(int, inicio.split(":"))
        h_out, m_out = map(int, fim.split(":"))
        delta = timedelta(hours=h_out, minutes=m_out) - timedelta(hours=h_in, minutes=m_in)
        horas_liquidas = delta.total_seconds() / 3600 - almoco
        return horas_liquidas

    # --- Cria o plano combinando registros existentes e preenchimento automático ---
    plano_completo = {}

    for data_str, info in registros.items():
        data_obj = str_para_data(data_str)
        
        # Só dias úteis até hoje
        if data_obj.weekday() <= 4 and data_obj <= hoje:
            if info["status"] != "vazio" or info["feriado"] or info["viagem"]:
                plano_completo[data_str] = info
            else:
                horas_turno = calcular_horas_liquidas(hora_entrada, hora_saida, tempo_almoco)
                if horas_turno > limite_horas:
                    ajuste_saida = datetime.strptime(hora_entrada, "%H:%M") + timedelta(hours=limite_horas + tempo_almoco)
                    hora_saida_ajustada = ajuste_saida.strftime("%H:%M")
                else:
                    hora_saida_ajustada = hora_saida
                
                plano_completo[data_str] = {
                    **info,
                    "in_1": hora_entrada,
                    "out_2": hora_saida_ajustada
                }

    # --- Salva o plano completo ---
    with open("plano_para_preenchimento.json", "w", encoding="utf-8") as f:
        json.dump(plano_completo, f, ensure_ascii=False, indent=2)

    print(f"✅ Plano completo gerado para {len(plano_completo)} dias úteis até hoje.")
    return plano_completo

if __name__ == "__main__":
    gerar_plano()
