import json
from datetime import datetime, timedelta, time
import pandas as pd
from pathlib import Path

# ----------------- CONFIG -----------------
ARQ_REGISTROS = "registros_mensais.json"
ARQ_PLANO = "plano_para_preenchimento.json"
ARQ_SF = r"C:\Users\brludas\Downloads\script pmg\script_horas\timesheet.xlsx"

HORA_MIN_INICIO = datetime.strptime("06:30", "%H:%M")       # antes disso => viagem
HORA_FLEX_END = datetime.strptime("10:00", "%H:%M")         # janela flexível até 10:00
HORA_PADRAO_IN_STR = "07:30"
HORA_PADRAO_OUT_STR = "16:54"
HORA_PADRAO_IN = datetime.strptime(HORA_PADRAO_IN_STR, "%H:%M")
HORA_PADRAO_OUT = datetime.strptime(HORA_PADRAO_OUT_STR, "%H:%M")

MINIMO_LIQUIDO = 8 + 24/60  # 8h24 -> 8.4 horas (líquido)
BRUTO_NECESSARIO = MINIMO_LIQUIDO + 1.0  # bruto a enviar para PMóvel (PMóvel subtrai 1h almoço)
LIMITE_DIA = 10  # limite da empresa (TAC se exceder)
MAX_DAY_TIME = datetime.strptime("23:59", "%H:%M")  # para cap de saída

# ------------- HELPERS -------------
def parse_hora(h):
    """Converte célula do Excel/string para datetime (apenas a parte time). Retorna None se inválido."""
    if pd.isna(h) or str(h).strip() == "":
        return None
    if isinstance(h, (pd.Timestamp, datetime)):
        return datetime.strptime(h.strftime("%H:%M"), "%H:%M")
    return datetime.strptime(str(h).strip(), "%H:%M")


def duracao_horas(blocos):
    """Soma duração (em horas) de blocos [(h1,h2), ...]."""
    total = 0.0
    for a, b in blocos:
        total += (b - a).total_seconds() / 3600
    return total


def cap_saida_no_dia(saida_dt):
    """Impede que a saída ultrapasse 23:59 do mesmo dia (fallback seguro)."""
    if saida_dt.time() > MAX_DAY_TIME.time():
        return MAX_DAY_TIME
    return saida_dt


def esta_dentro_padrao(h_inicio, h_fim):
    """Verifica se o bloco de labor está dentro do horário padrão (ou seja, começa e termina dentro)."""
    return (h_inicio >= HORA_PADRAO_IN and h_inicio <= HORA_PADRAO_OUT) and (h_fim >= HORA_PADRAO_IN and h_fim <= HORA_PADRAO_OUT)


# ------------- PROCESSAR DIA -------------
def processar_dia_com_sf(blocos_sf):
    """
    blocos_sf: lista de (tipo, h_in, h_out)
    retorna: in_1_str, out_1_str, status_str, descricao_str
    """
    blocos = []
    for t, h1, h2 in blocos_sf:
        if not (h1 and h2):
            continue
        tipo = str(t).strip().lower()
        if tipo == "labour":
            tipo = "labor"
        blocos.append((tipo, h1, h2))

    labors = [(h1, h2) for t, h1, h2 in blocos if t == "labor"]
    arrivals = [(h1, h2) for t, h1, h2 in blocos if t == "arrival"]
    deps = [(h1, h2) for t, h1, h2 in blocos if t == "departure"]

    total_labor = duracao_horas(labors)

    # ---------- CASO 1: LABOR >= MINIMO_LIQUIDO ----------
    if total_labor >= MINIMO_LIQUIDO and labors:
        entrada = min(h1 for h1, _ in labors)
        saida = max(h2 for _, h2 in labors)
        # não contar antes de 06:30 como trabalho
        if entrada < HORA_MIN_INICIO:
            entrada = HORA_MIN_INICIO
        status = "labor_suficiente"
        descricao = f"Labor {total_labor:.2f}h >= {MINIMO_LIQUIDO:.2f}h; enviar labor completo."
        if total_labor > LIMITE_DIA:
            status = "labor_suficiente_tac_required"
            descricao += " (labor > 10h: TAC requerido)"
        return entrada.strftime("%H:%M"), saida.strftime("%H:%M"), status, descricao

    # ---------- CASO 2: LABOR < MINIMO_LIQUIDO ----------
    # Se não houver labor e só arrivals/deps, e soma <= 10h -> aplicar padrão
    if not labors:
        total_arr_dep = duracao_horas(arrivals + deps)
        if total_arr_dep <= LIMITE_DIA:
            return HORA_PADRAO_IN_STR, HORA_PADRAO_OUT_STR, "padrao_manual", "Sem labor; preenchido com horário padrão."

    # montar blocos elegíveis (labor completo + porções de viagem dentro da janela flex)
    blocos_validos = list(labors)  # sempre incluir labor completo quando existente

    # construir lista de viagens elegíveis (porção entre 06:30 e 10:00)
    viagens_elegiveis = []
    for h1, h2 in arrivals + deps:
        if h2 <= HORA_MIN_INICIO:
            continue  # totalmente antes de 06:30 -> viagem, não elegível
        # parte relevante dentro da janela flexível (06:30 - 10:00)
        parte_inicio = max(h1, HORA_MIN_INICIO)
        parte_fim = min(h2, HORA_FLEX_END)
        if parte_fim > parte_inicio:
            viagens_elegiveis.append((parte_inicio, parte_fim))

    # Se temos viagens elegíveis, adiciona-as
    for b in viagens_elegiveis:
        blocos_validos.append(b)

    # --- Regra especial (o ajuste que você pediu) ---
    # Se há labor curto (<8h24) e NÃO há viagens elegíveis, e o labor acontece dentro do horário padrão,
    # é mais sensato enviar o horário padrão 07:30-16:54 (evita ampliar para noite).
    if labors and total_labor < MINIMO_LIQUIDO and not viagens_elegiveis:
        # verificar se o labor está dentro do período padrão
        primeiro_inicio_labor = min(h1 for h1, _ in labors)
        ultimo_fim_labor = max(h2 for _, h2 in labors)
        if esta_dentro_padrao(primeiro_inicio_labor, ultimo_fim_labor):
            return HORA_PADRAO_IN_STR, HORA_PADRAO_OUT_STR, "padrao_manual_from_labor", "Labor curto sem viagens elegíveis; aplicado horário padrão."

    # se ainda não há blocos válidos (rare), enviar mínimo bruto a partir de 06:30
    if not blocos_validos:
        entrada = HORA_MIN_INICIO
        saida = entrada + timedelta(hours=BRUTO_NECESSARIO)
        saida = cap_saida_no_dia(saida)
        status = "labor_insuficiente_completado"
        descricao = f"Sem blocos válidos; enviado bruto {BRUTO_NECESSARIO:.2f}h."
        return entrada.strftime("%H:%M"), saida.strftime("%H:%M"), status, descricao

    # calcular entrada como menor início entre blocos_validos
    entrada = min(h1 for h1, _ in blocos_validos)
    # garantir entrada nunca antes de 06:30
    if entrada < HORA_MIN_INICIO:
        entrada = HORA_MIN_INICIO

    # determinar último fim real de labor (preferir fim do labor, pois labor deve ser enviado completo)
    if labors:
        ultimo_fim_labor = max(h2 for _, h2 in labors)
    else:
        ultimo_fim_labor = max(h2 for _, h2 in blocos_validos)

    # saída mínima para garantir bruto
    saida_por_bruto = entrada + timedelta(hours=BRUTO_NECESSARIO)

    # saída final é o máximo entre fim real do labor (se existir) e saída por bruto
    saida = max(ultimo_fim_labor, saida_por_bruto)

    # cap para não extrapolar dia
    saida = cap_saida_no_dia(saida)

    # montar status/descrição
    total_envio = (saida - entrada).total_seconds() / 3600
    status = "labor_insuficiente_completado"
    descricao = (f"Labor {total_labor:.2f}h < {MINIMO_LIQUIDO:.2f}h; "
                 f"complementado com viagens elegíveis. Enviado bruto {total_envio:.2f}h "
                 f"de {entrada.strftime('%H:%M')} até {saida.strftime('%H:%M')}.")

    # marcar tac se ultrapassar limite diário
    if total_envio > LIMITE_DIA:
        status += "_tac_required"
        descricao += " (TAC requerido, >10h)"

    return entrada.strftime("%H:%M"), saida.strftime("%H:%M"), status, descricao


# ------------- FUNÇÃO PRINCIPAL -------------
def gerar_plano():
    if not Path(ARQ_REGISTROS).exists():
        print("Arquivo registros_mensais.json não encontrado.")
        return

    with open(ARQ_REGISTROS, "r", encoding="utf-8") as f:
        registros = json.load(f)

    hoje = datetime.today().date()

    # ler Salesforce (timesheet)
    if Path(ARQ_SF).exists():
        df = pd.read_excel(ARQ_SF, engine="openpyxl")
    else:
        df = pd.DataFrame(columns=["Data", "Hora início", "Hora fim", "Tipo"])

    # normalizar colunas
    df["Data"] = pd.to_datetime(df.get("Data", pd.NaT), dayfirst=True, errors="coerce")
    df["Hora início"] = df.get("Hora início", "")
    df["Hora fim"] = df.get("Hora fim", "")
    df["Tipo"] = df.get("Tipo", "")

    # agrupar por data
    atendimentos_sf = {}
    for _, row in df.iterrows():
        if pd.isna(row["Data"]):
            continue
        data_str = row["Data"].strftime("%d/%m/%Y")
        tipo = str(row.get("Tipo", "")).strip().lower()
        if tipo == "labour":
            tipo = "labor"
        h_in = parse_hora(row.get("Hora início"))
        h_out = parse_hora(row.get("Hora fim"))
        if h_in is None or h_out is None:
            continue
        atendimentos_sf.setdefault(data_str, []).append((tipo, h_in, h_out))

    plano_final = {}

    for data_str, info in registros.items():
        data_obj = datetime.strptime(data_str, "%d/%m/%Y").date()

        # ignorar finais de semana e datas futuras
        if data_obj.weekday() > 4 or data_obj > hoje:
            continue

        if data_str in atendimentos_sf:
            in1, out1, status, desc = processar_dia_com_sf(atendimentos_sf[data_str])
            novo = dict(info)
            novo.update({
                "in_1": in1,
                "out_1": out1,
                "in_2": None,
                "out_2": None,
                "in_3": None,
                "out_3": None,
                "in_4": None,
                "out_4": None,
                "status": status,
                "descricao_status": desc
            })
            plano_final[data_str] = novo
        else:
            # aplica padrão mesmo que dia não esteja no SF (preserva todos os dias)
            if info.get("status") == "vazio" and not info.get("feriado") and not info.get("viagem"):
                novo = dict(info)
                novo.update({
                    "in_1": HORA_PADRAO_IN_STR,
                    "out_1": HORA_PADRAO_OUT_STR,
                    "in_2": None,
                    "out_2": None,
                    "in_3": None,
                    "out_3": None,
                    "in_4": None,
                    "out_4": None,
                    "status": "padrao_manual",
                    "descricao_status": "Preenchido com horário padrão"
                })
                plano_final[data_str] = novo
            else:
                plano_final[data_str] = info

    # salvar plano final
    with open(ARQ_PLANO, "w", encoding="utf-8") as f:
        json.dump(plano_final, f, ensure_ascii=False, indent=2)

    print(f"✅ Plano integrado gerado ({len(plano_final)} dias). Salvo em: {ARQ_PLANO}")


if __name__ == "__main__":
    gerar_plano()
