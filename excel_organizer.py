# emailtoexcel_improved.py
# Requer: pip install pandas openpyxl
# Uso: python emailtoexcel_improved.py

import pandas as pd
import re
from pathlib import Path

# --- Ajuste só estes caminhos ---
INPUT = "C:\\Users\\brludas\\Downloads\\script pmg\\script_horas\\tabela_final_organizada.xlsx"
OUTPUT = "timesheet.xlsx"
# ---------------------------------

def limpar_texto(s):
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

# Regexes
date_re = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})")
datetime_re = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}:\d{2})")
time_re = re.compile(r"(\d{1,2}:\d{2})")
duration_re = re.compile(r"(?<!\d)(\d{1,3}[.,]\d{1,2}|\d{1,2}[.,]\d{1,2})(?!\d)")
ot_re = re.compile(r"(?<!\d)(\d{5,9})(?!\d)")

# Tipos conhecidos (extenda se precisar)
types_list = ["arrival", "departure", "labour", "travel", "work", "break"]
types_re = re.compile(r"\b(" + "|".join([re.escape(t) for t in types_list]) + r")\b", re.IGNORECASE)

# --- Detectar cliente automaticamente ---
# Heurística:
# cliente = texto que parece nome próprio, curto, não é data, hora, duração, tipo ou OT
cliente_re = re.compile(
    r"^(?:[A-Z]{2,10}|[A-ZÁÉÍÓÚÂÊÔÃÕ][\w\-\.' ]{2,})$"
)


def processar_arquivo(input_path, output_path):
    print("Lendo:", input_path)
    df = pd.read_excel(input_path, engine="openpyxl")

    # Normaliza texto em todas as células
    df_clean = df.fillna("").astype(str).applymap(limpar_texto)

    registros = []
    current_date = None

    for idx, row in df_clean.iterrows():
        joined = " | ".join([c for c in row if c])
        if not joined:
            continue

        # pular linhas de cabeçalho detectáveis
        low = joined.lower()
        if "data de início" in low or "hora de início" in low or "duration" in low or "data de início da planilha" in low:
            continue

        rec = {"Data": None, "Hora início": None, "Hora fim": None, "Duração": None,
               "Tipo": None, "Cliente": None, "OT": None, "Descrição": None,
               "raw": joined, "orig_row": idx}

        horas = []
        descr = []

        for cell in row:
            if not cell:
                continue
            cell_text = str(cell).strip()
            cell_low = cell_text.lower()

            # datetime (dd/mm/yyyy HH:MM)
            dtm = datetime_re.search(cell_text)
            if dtm:
                current_date = dtm.group(1)
                horas.append(dtm.group(2))
                cell_text = datetime_re.sub("", cell_text).strip()

            # date
            d = date_re.search(cell_text)
            if d:
                current_date = d.group(1)
                cell_text = date_re.sub("", cell_text).strip()

            # times
            times = time_re.findall(cell_text)
            if times:
                for t in times:
                    horas.append(t)
                cell_text = time_re.sub("", cell_text).strip()

            # duração
            dur = duration_re.findall(cell_text)
            if dur and rec["Duração"] is None:
                # escolhe primeiro que pareça duração
                for dm in dur:
                    # evita pegar coisas estranhas
                    cand = dm.replace(",", ".")
                    try:
                        v = float(cand)
                        rec["Duração"] = v
                        # remove apenas a primeira ocorrência
                        cell_text = duration_re.sub("", cell_text, count=1).strip()
                        break
                    except:
                        continue

            # tipo
            tp = types_re.search(cell_text)
            if tp and not rec["Tipo"]:
                rec["Tipo"] = tp.group(1).capitalize()
                cell_text = types_re.sub("", cell_text).strip()

            # OT
            otm = ot_re.search(cell_text)
            if otm and not rec["OT"]:
                rec["OT"] = otm.group(1)
                cell_text = ot_re.sub("", cell_text).strip()

            # Detecta cliente automaticamente:
            if not rec["Cliente"]:
                # candidato = texto curto, começa com maiúscula, não tem hora/data, não é tipo
                if (
                    2 <= len(cell_text.split()) <= 5 and
                    not date_re.search(cell_text) and
                    not datetime_re.search(cell_text) and
                    not time_re.search(cell_text) and
                    not duration_re.search(cell_text) and
                    not types_re.search(cell_text.lower()) and
                    not ot_re.search(cell_text) and
                    cliente_re.match(cell_text)
                ):
                    rec["Cliente"] = cell_text
                    cell_text = ""

            # leftover -> parte da descrição
            if cell_text:
                descr.append(cell_text)

        # carry-forward de data
        if current_date:
            rec["Data"] = current_date

        # aloca horas heurística
        if horas:
            if len(horas) == 1:
                rec["Hora início"] = horas[0]
            else:
                rec["Hora início"] = horas[0]
                rec["Hora fim"] = horas[1]

        if descr:
            rec["Descrição"] = " | ".join(descr).strip()

        # Se não encontrou nada significativo, pule (evita lixos)
        meaningful = any([rec[k] for k in ["Data","Hora início","Hora fim","Duração","Tipo","OT","Cliente","Descrição"]])
        if meaningful:
            registros.append(rec)

    # Cria DataFrame final
    df_final = pd.DataFrame(registros, columns=["Data","Hora início","Hora fim","Duração","Tipo","Cliente","OT","Descrição","raw","orig_row"])

    # Remove duplicados óbvios (mesma data, hora e OT)
    df_final = df_final.drop_duplicates(subset=["Data","Hora início","OT","Tipo"], keep="first").reset_index(drop=True)

    # Salva
    df_final.to_excel(output_path, index=False, engine="openpyxl")
    print("Salvo em:", output_path)
    return df_final

if __name__ == "__main__":
    df_out = processar_arquivo(INPUT, OUTPUT)
    print("Exemplo (top 10):")
    print(df_out.head(10))
