# gerar_plano.py
"""
Gerador de plano PMóvel - lado A (trabalho) + lado B (viagens sobrantes).
Regras principais (resumo):
- labor >= 8h24 -> enviar labor completo; viagens NÃO usadas.
- labor < 8h24 -> complementar usando arrivals (antes do 1º labor) e departures (após o último labor),
  consumindo apenas as partes necessárias até atingir o bruto desejado.
- Lado B recebe apenas as sobras NÃO consumidas.
- Ajuste -1 minuto só em arrivals que serão enviados ao PMóvel (lado B) quando seu fim >= 07:30.
- Dias sem SF e vazios -> horário padrão 07:30-16:54.
"""

import json
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
from copy import deepcopy

# ----------------- CONFIG -----------------
ARQ_REGISTROS = "registros_mensais.json"
ARQ_PLANO = "plano_para_preenchimento.json"
ARQ_SF = r"C:\Users\brludas\Downloads\script pmg\script_horas\timesheet.xlsx"

HORA_MIN_INICIO = datetime.strptime("06:30", "%H:%M")
HORA_FLEX_END = datetime.strptime("10:00", "%H:%M")
HORA_PADRAO_IN_STR = "07:30"
HORA_PADRAO_OUT_STR = "16:54"
HORA_PADRAO_IN = datetime.strptime(HORA_PADRAO_IN_STR, "%H:%M")
HORA_PADRAO_OUT = datetime.strptime(HORA_PADRAO_OUT_STR, "%H:%M")

MINIMO_LIQUIDO = 8 + 24/60   # 8h24 = 8.4 horas (líquido)
BRUTO_NECESSARIO = MINIMO_LIQUIDO + 1.0
LIMITE_DIA = 10
MAX_DAY_TIME = datetime.strptime("23:59", "%H:%M")
PMOVEL_BLOCK_TRIM_TIME = datetime.strptime("07:30", "%H:%M").time()  # arrivals ending >= this need -1min when sent to B

# ------------- HELPERS -------------
def parse_hora(h):
    if pd.isna(h) or str(h).strip() == "":
        return None
    if isinstance(h, (pd.Timestamp, datetime)):
        return datetime.strptime(h.strftime("%H:%M"), "%H:%M")
    return datetime.strptime(str(h).strip(), "%H:%M")

def duracao_horas(blocos):
    total = 0.0
    for a, b in blocos:
        total += (b - a).total_seconds() / 3600
    return total

def cap_saida_no_dia(saida_dt):
    if not isinstance(saida_dt, datetime):
        return saida_dt
    if saida_dt.time() > MAX_DAY_TIME.time():
        return datetime.combine(saida_dt.date(), MAX_DAY_TIME.time())
    return saida_dt

def subtrair_intervalo(original, usado):
    """
    original: (o1,o2)
    usado: (u1,u2) guaranteed subset of original
    retorna lista de intervalos leftovers (0,1 ou 2)
    """
    o1, o2 = original
    u1, u2 = usado
    restos = []
    if u1 > o1:
        restos.append((o1, min(u1, o2)))
    if u2 < o2:
        restos.append((max(u2, o1), o2))
    return restos

def ajustar_arrival_para_pmovel(h1, h2):
    """Se fim >= 07:30, aplica -1 minuto em ambos os lados; caso contrário retorna inalterado."""
    if h2.time() >= PMOVEL_BLOCK_TRIM_TIME:
        return (h1 - timedelta(minutes=1), h2 - timedelta(minutes=1))
    return (h1, h2)

# ------------- ALLOCADOR DE SUPLEMENTO -------------
def alocar_complementos(labors, arrivals, departures):
    """
    Decide quais partes de arrivals/departures serão USADAS para compor o bloco de trabalho.
    Retorna:
      - entrada (datetime), saida (datetime), status, descricao,
      - used_arrivals (list of (h1,h2)),
      - used_departures (list of (h1,h2)),
      - leftover_arrivals (list of (h1,h2)),
      - leftover_departures (list of (h1,h2))
    """
    # deep copies para manipular leftovers
    leftover_arr = deepcopy(arrivals)
    leftover_dep = deepcopy(departures)

    total_labor = duracao_horas(labors)

    # Caso clássico: labor existe porém é insuficiente e não há viagens.
# PMóvel deve receber o padrão.
    if 0 < total_labor < MINIMO_LIQUIDO and not arrivals and not departures:
        entrada = HORA_PADRAO_IN
        saida = HORA_PADRAO_OUT
        return (
            entrada,
            saida,
            "padrao_por_labor_curto",
            f"Labor {total_labor:.2f}h insuficiente e sem viagens; enviado padrão.",
            [],
            [],
            arrivals[:],
            departures[:]
        )


    # Caso labor suficiente: usamos todo o labor e não tocamos viagens
    if total_labor >= MINIMO_LIQUIDO and labors:
        entrada = min(h1 for h1,_ in labors)
        saida = max(h2 for _,h2 in labors)
        if entrada < HORA_MIN_INICIO:
            entrada = HORA_MIN_INICIO
        status = "labor_suficiente"
        descricao = f"Labor {total_labor:.2f}h >= {MINIMO_LIQUIDO:.2f}h; enviar labor completo."
        if total_labor > LIMITE_DIA:
            status = "labor_suficiente_tac_required"
            descricao += " (labor > 10h: TAC requerido)"
        # no arrivals/departures used
        return entrada, saida, status, descricao, [], [], leftover_arr, leftover_dep

    # Caso sem labor (somente viagens):
    if not labors:
        total_arr_dep = duracao_horas(arrivals + departures)
        if total_arr_dep <= LIMITE_DIA:
            # aplicar padrão (não consumir viagens) — já acordado
            entrada = HORA_PADRAO_IN
            saida = HORA_PADRAO_OUT
            return entrada, saida, "padrao_manual", "Sem labor; preenchido padrão.", [], [], leftover_arr, leftover_dep
        # se precisa construir, consideraremos arrivals+departures como disponíveis

    # Precisamos compor bloco que contenha labors completos (se existirem) e
    # usar arrivals antes do primeiro labor e departures depois do último labor, na ordem:
    # 1) arrivals (do mais cedo ao mais tarde) — só a parte antes do primeiro labor
    # 2) labors (todos completos)
    # 3) departures (do mais cedo ao mais tarde) — só a parte após o último labor
    # o objetivo: garantir que (saida - entrada) >= BRUTO_NECESSARIO (bruto)

    # calcula entrada base e fim base
    if labors:
        first_labor_start = min(h1 for h1,_ in labors)
        last_labor_end = max(h2 for _,h2 in labors)
    else:
        # sem labor, consideramos artificialmente first/last como None e usaremos arrivals/deps
        first_labor_start = None
        last_labor_end = None

    # disponíveis: arrivals antes do first_labor_start (considerando parte após 06:30)
    available_arrivals = []
    for a1,a2 in arrivals:
        # considerar apenas parte > 06:30
        if a2 <= HORA_MIN_INICIO:
            continue
        # parcela antes do labor start (se existir), senão parcela inteira antes de flex_end
        start = max(a1, HORA_MIN_INICIO)
        end = a2 if first_labor_start is None else min(a2, first_labor_start)
        if end > start:
            available_arrivals.append((a1, a2, start, end))  # keep original and usable part

    # disponíveis: departures after last_labor_end (part after last_labor_end)
    available_departures = []
    for d1,d2 in departures:
        if d2 <= HORA_MIN_INICIO:
            continue
        start = d1 if last_labor_end is None else max(d1, last_labor_end)
        end = d2
        if end > start:
            available_departures.append((d1, d2, start, end))

    # compute current bruto span if labors exist:
    if labors:
        # entry is either first labor start or earlier if we include arrival
        # initially set entrada to first labor start (clipped)
        entrada_candidate = max(first_labor_start, HORA_MIN_INICIO) if first_labor_start else HORA_MIN_INICIO
        fim_labor = last_labor_end if last_labor_end else entrada_candidate
    else:
        # without labor, choose earliest available arrival start or earliest dep start
        all_starts = []
        for a1,a2 in arrivals:
            if a2 > HORA_MIN_INICIO:
                all_starts.append(max(a1,HORA_MIN_INICIO))
        for d1,d2 in departures:
            if d2 > HORA_MIN_INICIO:
                all_starts.append(max(d1,HORA_MIN_INICIO))
        entrada_candidate = min(all_starts) if all_starts else HORA_MIN_INICIO
        fim_labor = entrada_candidate

    # compute initial bruto from current blocks (labors plus any arrivals we will include)
    # We'll greedily include arrivals (earliest usable part) until needed, then include departures as needed.
    used_arrivals = []
    used_departures = []

    # start with labor span
    if labors:
        current_start = entrada_candidate
        current_end = fim_labor
    else:
        current_start = entrada_candidate
        current_end = entrada_candidate

    current_bruto = (current_end - current_start).total_seconds()/3600

    # need to reach at least BRUTO_NECESSARIO
    # first, use arrivals in chronological order (they extend the start backward)
    # Each arrival usable portion (start_use, end_use) is before first_labor_start
    # Using an arrival extends current_start earlier to min(current_start, start_use) and increases bruto accordingly.
    for orig_a1, orig_a2, use_start, use_end in sorted(available_arrivals, key=lambda x: x[2]):
        if current_bruto >= BRUTO_NECESSARIO:
            break
        # determine how much of this usable part is needed
        # if we include all usable part, new_start = min(current_start, use_start)
        new_start = min(current_start, use_start)
        added = (current_end - new_start).total_seconds()/3600 - current_bruto
        # if added <= 0 then nothing new; else include
        if added > 0:
            # include only required portion from the usable part: if added >= (use_end-use_start) then include whole usable
            usable_len = (use_end - use_start).total_seconds()/3600
            need = BRUTO_NECESSARIO - current_bruto
            # if need >= usable_len -> take whole usable portion (use_start..use_end); else take partial from the end of usable portion
            if need >= usable_len:
                take_start = use_start
            else:
                # take the last `need` hours from usable segment, so the taken segment ends at use_end and starts at use_end - need
                take_start = use_end - timedelta(hours=need)
                if take_start < use_start:
                    take_start = use_start
            take_end = use_end
            used_arrivals.append((take_start, take_end))
            current_start = min(current_start, take_start)
            current_bruto = (current_end - current_start).total_seconds()/3600

    # then, if still not enough, use departures after last labor end (they extend end forward)
    for orig_d1, orig_d2, use_start, use_end in sorted(available_departures, key=lambda x: x[2]):
        if current_bruto >= BRUTO_NECESSARIO:
            break
        # if current_end < use_start, including departure will extend end to at least use_start, then to use_end
        # compute new end if we take part
        # determine portion length available:
        usable_len = (use_end - use_start).total_seconds()/3600
        need = BRUTO_NECESSARIO - current_bruto
        if need <= 0:
            break
        if need >= usable_len:
            take_start = use_start
            take_end = use_end
        else:
            take_start = use_start
            take_end = use_start + timedelta(hours=need)
        used_departures.append((take_start, take_end))
        current_end = max(current_end, take_end)
        current_bruto = (current_end - current_start).total_seconds()/3600

    # if still not enough (rare), extend the end artificially (we will not mark this as used_departure)
    if current_bruto < BRUTO_NECESSARIO:
        need = BRUTO_NECESSARIO - current_bruto
        current_end = current_end + timedelta(hours=need)
        current_end = cap_saida_no_dia(current_end)
        current_bruto = (current_end - current_start).total_seconds()/3600

    # now we've determined which parts were used: used_arrivals, used_departures
    # compute leftovers by subtracting used parts from original arrivals/departures

    def subtract_used_from_originals(originals, used_list):
        leftovers = []
        used = sorted(used_list, key=lambda x: x[0])
        for orig in originals:
            o1,o2 = orig
            to_process = [(o1,o2)]
            for u1,u2 in used:
                new_proc = []
                for seg in to_process:
                    s1,s2 = seg
                    # if no overlap, keep seg
                    if u2 <= s1 or u1 >= s2:
                        new_proc.append(seg)
                    else:
                        # subtract overlap
                        parts = subtrair_intervalo((s1,s2), (max(s1,u1), min(s2,u2)))
                        for p in parts:
                            new_proc.append(p)
                to_process = new_proc
            for p in to_process:
                # only keep positive-length
                if p[1] > p[0]:
                    leftovers.append(p)
        return leftovers

    leftover_arrivals_final = subtract_used_from_originals(arrivals, used_arrivals)
    leftover_departures_final = subtract_used_from_originals(departures, used_departures)

    # finalize entrada/saida: entrada = current_start, saida = current_end (cap day)
    entrada = current_start
    saida = cap_saida_no_dia(current_end)

    status = "labor_insuficiente_completado"
    dur_bruto = (saida - entrada).total_seconds() / 3600
    descricao = (
        f"Labor {total_labor:.2f}h < {MINIMO_LIQUIDO:.2f}h; "
        f"completado. Enviado {dur_bruto:.2f}h."
    )

    if (saida - entrada).total_seconds()/3600 > LIMITE_DIA:
        status += "_tac_required"
        descricao += " (TAC requerido, >10h)"

    return entrada, saida, status, descricao, used_arrivals, used_departures, leftover_arrivals_final, leftover_departures_final

# ------------- FUNÇÃO PRINCIPAL -------------
def gerar_plano():
    if not Path(ARQ_REGISTROS).exists():
        print("Arquivo registros_mensais.json não encontrado.")
        return

    with open(ARQ_REGISTROS, "r", encoding="utf-8") as f:
        registros = json.load(f)

    hoje = datetime.today().date()

    # ler Salesforce
    if Path(ARQ_SF).exists():
        df = pd.read_excel(ARQ_SF, engine="openpyxl")
    else:
        df = pd.DataFrame(columns=["Data","Hora início","Hora fim","Tipo"])

    df["Data"] = pd.to_datetime(df.get("Data", pd.NaT), dayfirst=True, errors="coerce")
    df["Hora início"] = df.get("Hora início", "")
    df["Hora fim"] = df.get("Hora fim", "")
    df["Tipo"] = df.get("Tipo", "")

    atendimentos_sf = {}
    for _, row in df.iterrows():
        if pd.isna(row["Data"]):
            continue
        d = row["Data"].strftime("%d/%m/%Y")
        tipo = str(row.get("Tipo","")).strip().lower()
        if tipo == "labour":
            tipo = "labor"
        h1 = parse_hora(row.get("Hora início"))
        h2 = parse_hora(row.get("Hora fim"))
        if h1 and h2:
            atendimentos_sf.setdefault(d, []).append((tipo, h1, h2))

    plano_final = {}

    for data_str, info in registros.items():
        data_obj = datetime.strptime(data_str, "%d/%m/%Y").date()
        if data_obj.weekday() > 4 or data_obj > hoje:
            continue

        # default values for travel arrays
        viagens_arr_json = []
        viagens_dep_json = []

        if data_str in atendimentos_sf:
            # split blocks into types
            blocos = atendimentos_sf[data_str]
            labors = [(h1,h2) for t,h1,h2 in blocos if t=="labor"]
            arrivals = [(h1,h2) for t,h1,h2 in blocos if t=="arrival"]
            departures = [(h1,h2) for t,h1,h2 in blocos if t=="departure"]

            entrada, saida, status, desc, used_arr, used_dep, leftover_arr, leftover_dep = alocar_complementos(labors, arrivals, departures)

            # prepare viagens (lado B) from leftovers only
            # apply adjustment -1min only on arrivals leftovers whose end >= 07:30
            for h1,h2 in leftover_arr:
                adj1, adj2 = ajustar_arrival_para_pmovel(h1,h2)
                viagens_arr_json.append({"inicio": adj1.strftime("%H:%M"), "fim": adj2.strftime("%H:%M")})

            # departures leftovers go as-is
            for h1,h2 in leftover_dep:
                viagens_dep_json.append({"inicio": h1.strftime("%H:%M"), "fim": h2.strftime("%H:%M")})

            novo = dict(info)
            novo.update({
                "in_1": entrada.strftime("%H:%M"),
                "out_1": saida.strftime("%H:%M"),
                "in_2": None, "out_2": None,
                "in_3": None, "out_3": None,
                "in_4": None, "out_4": None,
                "status": status,
                "descricao_status": desc,
                "viagens_arrival": viagens_arr_json,
                "viagens_departure": viagens_dep_json
            })
            plano_final[data_str] = novo

        else:
            # dia sem SF — preencher padrão se vazio e não feriado/viagem
            if info.get("status") == "vazio" and not info.get("feriado") and not info.get("viagem"):
                novo = dict(info)
                novo.update({
                    "in_1": HORA_PADRAO_IN_STR,
                    "out_1": HORA_PADRAO_OUT_STR,
                    "in_2": None, "out_2": None,
                    "in_3": None, "out_3": None,
                    "in_4": None, "out_4": None,
                    "status": "padrao_manual",
                    "descricao_status": "Preenchido com horário padrão",
                    "viagens_arrival": [],
                    "viagens_departure": []
                })
                plano_final[data_str] = novo
            else:
                info["viagens_arrival"] = []
                info["viagens_departure"] = []
                plano_final[data_str] = info

    # salvar JSON final
    with open(ARQ_PLANO, "w", encoding="utf-8") as f:
        json.dump(plano_final, f, ensure_ascii=False, indent=2)

    print("✅ Plano completo gerado: lado A (trabalho) + lado B (viagens sobrantes).")

if __name__ == "__main__":
    gerar_plano()
