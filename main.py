import io
import json
import math
import statistics
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(
    page_title="Senior Technical Consultant Selector",
    page_icon="🏅",
    layout="wide",
)

# =========================================================
# CONFIGURACIÓN DEL MODELO
# =========================================================

CRITICAL_INDICATORS = {
    "Nutrición en general": 1.00,
    "Conocimiento de productos": 1.00,
    "Patología metabólica": 1.10,
    "Aditivos alternativos": 1.00,
    "Bioseguridad": 1.00,
    "Arranques": 1.10,
    "Animales en producción": 1.10,
    "Tratamiento de datos": 1.00,
    "Informes": 0.95,
    "Ingles": 0.90,
    "Manejo herramientas/programas": 0.95,
}

TRANSFER_INDICATORS = [
    "Conocimiento de productos",
    "Tratamiento de datos",
    "Tratamiento de textos",
    "Informes",
    "Ingles",
    "Manejo herramientas/programas",
]

RANKING_WEIGHTS = {
    "global_performance": 0.35,
    "balance": 0.20,
    "critical": 0.20,
    "transfer": 0.15,
    "team_advantage": 0.10,
}

CRITERIA_TABLE = pd.DataFrame(
    [
        {
            "Criterio": "Rendimiento técnico global",
            "Peso": "35%",
            "Qué mide": "Nivel técnico total del candidato dentro de la propia herramienta.",
            "Cómo se calcula": "Combinación de la comparación global vs objetivo y de la comparación global vs máximo de referencia.",
        },
        {
            "Criterio": "Equilibrio entre los 4 troncos",
            "Peso": "20%",
            "Qué mide": "Si el perfil es completo y consistente o está demasiado descompensado.",
            "Cómo se calcula": "Media de los 4 troncos vs objetivo + refuerzo del tronco más débil para penalizar desequilibrios.",
        },
        {
            "Criterio": "Indicadores críticos para senior",
            "Peso": "20%",
            "Qué mide": "Fortaleza en las materias más relevantes para actuar como referente senior.",
            "Cómo se calcula": "Promedio ponderado de indicadores críticos vs objetivo y vs máximo, con penalización por carencias graves.",
        },
        {
            "Criterio": "Capacidad de transferencia y formación",
            "Peso": "15%",
            "Qué mide": "Potencial para formar, estructurar criterio y ayudar a crecer al resto del equipo.",
            "Cómo se calcula": "Subconjunto de indicadores ligados a comunicación, informes, datos, herramientas e inglés.",
        },
        {
            "Criterio": "Ventaja respecto a la media del equipo",
            "Peso": "10%",
            "Qué mide": "Cuánto destaca realmente el candidato frente al estándar medio del equipo.",
            "Cómo se calcula": "Comparación global vs BBDD + porcentaje de indicadores por encima de la media BBDD.",
        },
    ]
)

CRITERIA_TEXT = """
**Regla clave de justicia:** la app no elige automáticamente a quien tiene la mayor nota global.
Elige al perfil más adecuado para actuar como **consultor senior**, es decir, un técnico sólido,
equilibrado, fuerte en los indicadores críticos y con capacidad de referencia y de formación.

**Antes del ranking final hay un filtro mínimo de elegibilidad senior.**
Un candidato solo se considera “Apto ahora” si cumple los cuatro requisitos:
1. Está **por encima del objetivo global**.
2. No tiene ningún tronco **claramente débil**.
3. No acumula carencias graves en los **indicadores críticos**.
4. Tiene una base suficiente en **transferencia/formación**.

Si nadie cumple ese estándar, la app puede concluir que **todavía no hay un senior claro**.
"""


# =========================================================
# UTILIDADES
# =========================================================

def safe_float(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None


def ratio_to_score(value, stretch=1.20):
    value = safe_float(value)
    if value is None or math.isnan(value):
        return 0.0
    value = max(0.0, min(value, stretch))
    return (value / stretch) * 100.0


def pct(value):
    value = safe_float(value)
    if value is None:
        return None
    return round(value * 100.0, 1)


def format_pct(value):
    value = pct(value)
    return "-" if value is None else f"{value:.1f}%"


def strongest_and_weakest_trunk(trunks_df: pd.DataFrame):
    if trunks_df.empty:
        return "-", "-"
    tmp = trunks_df.copy().dropna(subset=["vs_goal"])
    if tmp.empty:
        return "-", "-"
    strongest = tmp.sort_values("vs_goal", ascending=False).iloc[0]["tronco"]
    weakest = tmp.sort_values("vs_goal", ascending=True).iloc[0]["tronco"]
    return strongest, weakest


def build_indicator_frame(ref_ws, eval_ws) -> pd.DataFrame:
    rows = []
    current_area = None

    for ref_row, eval_row in zip(range(4, 29), range(9, 34)):
        area = ref_ws[f"B{ref_row}"].value
        if area:
            current_area = str(area).strip()

        indicator = str(ref_ws[f"C{ref_row}"].value).strip()
        weight = safe_float(ref_ws[f"D{ref_row}"].value) or 0.0
        objective_raw = safe_float(ref_ws[f"E{ref_row}"].value) or 0.0
        objective_weighted = safe_float(ref_ws[f"F{ref_row}"].value)
        max_weighted = safe_float(ref_ws[f"G{ref_row}"].value)
        bbdd_raw = safe_float(ref_ws[f"H{ref_row}"].value)
        bbdd_weighted = safe_float(ref_ws[f"I{ref_row}"].value)

        raw_score = safe_float(eval_ws[f"D{eval_row}"].value)
        weighted_score = None if raw_score is None else raw_score * weight

        if objective_weighted is None:
            objective_weighted = weight * objective_raw
        if max_weighted is None:
            max_weighted = weight * 4
        if bbdd_weighted is None and bbdd_raw is not None:
            bbdd_weighted = weight * bbdd_raw

        vs_goal = None if not objective_weighted else weighted_score / objective_weighted if weighted_score is not None else None
        vs_max = None if not max_weighted else weighted_score / max_weighted if weighted_score is not None else None
        vs_bbdd = None if not bbdd_weighted else weighted_score / bbdd_weighted if weighted_score is not None else None

        rows.append(
            {
                "tronco": current_area,
                "indicator": indicator,
                "weight": weight,
                "score_raw": raw_score,
                "score_weighted": weighted_score,
                "objective_raw": objective_raw,
                "objective_weighted": objective_weighted,
                "max_weighted": max_weighted,
                "bbdd_raw": bbdd_raw,
                "bbdd_weighted": bbdd_weighted,
                "vs_goal": vs_goal,
                "vs_max": vs_max,
                "vs_bbdd": vs_bbdd,
            }
        )

    return pd.DataFrame(rows)


def summarise_trunks(indicators_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for trunk, grp in indicators_df.groupby("tronco", dropna=False):
        rows.append(
            {
                "tronco": trunk,
                "score_raw_total": grp["score_raw"].sum(skipna=True),
                "score_weighted_total": grp["score_weighted"].sum(skipna=True),
                "objective_weighted_total": grp["objective_weighted"].sum(skipna=True),
                "max_weighted_total": grp["max_weighted"].sum(skipna=True),
                "bbdd_weighted_total": grp["bbdd_weighted"].sum(skipna=True),
                "avg_raw": grp["score_raw"].mean(skipna=True),
                "vs_goal": (grp["score_weighted"].sum(skipna=True) / grp["objective_weighted"].sum(skipna=True))
                if grp["objective_weighted"].sum(skipna=True) else None,
                "vs_max": (grp["score_weighted"].sum(skipna=True) / grp["max_weighted"].sum(skipna=True))
                if grp["max_weighted"].sum(skipna=True) else None,
                "vs_bbdd": (grp["score_weighted"].sum(skipna=True) / grp["bbdd_weighted"].sum(skipna=True))
                if grp["bbdd_weighted"].sum(skipna=True) else None,
            }
        )
    return pd.DataFrame(rows)


def summarise_global(indicators_df: pd.DataFrame) -> dict:
    score_raw_total = indicators_df["score_raw"].sum(skipna=True)
    score_weighted_total = indicators_df["score_weighted"].sum(skipna=True)
    objective_weighted_total = indicators_df["objective_weighted"].sum(skipna=True)
    max_weighted_total = indicators_df["max_weighted"].sum(skipna=True)
    bbdd_weighted_total = indicators_df["bbdd_weighted"].sum(skipna=True)
    avg_raw = indicators_df["score_raw"].mean(skipna=True)
    bbdd_raw_avg = indicators_df["bbdd_raw"].mean(skipna=True)

    return {
        "score_raw_total": score_raw_total,
        "score_weighted_total": score_weighted_total,
        "objective_weighted_total": objective_weighted_total,
        "max_weighted_total": max_weighted_total,
        "bbdd_weighted_total": bbdd_weighted_total,
        "avg_raw": avg_raw,
        "bbdd_raw_avg": bbdd_raw_avg,
        "vs_goal": (score_weighted_total / objective_weighted_total) if objective_weighted_total else None,
        "vs_max": (score_weighted_total / max_weighted_total) if max_weighted_total else None,
        "vs_bbdd": (score_weighted_total / bbdd_weighted_total) if bbdd_weighted_total else None,
    }


def infer_level(avg_raw: float) -> str:
    if avg_raw is None or math.isnan(avg_raw):
        return "-"
    if avg_raw < 1:
        return "BÁSICO"
    if avg_raw < 2:
        return "CONTROLA"
    if avg_raw < 3:
        return "SUPERA"
    if avg_raw < 4:
        return "CERTIFICADO"
    if avg_raw < 5:
        return "EXCELENTE"
    if avg_raw < 6:
        return "MASTER"
    return "MÁXIMO"


def parse_candidate(uploaded_file):
    suffix = Path(uploaded_file.name).suffix or ".xlsm"

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        temp_path = tmp.name

    wb = load_workbook(temp_path, data_only=True, keep_vba=True)
    required = {"REFERENCIAS", "EVALUACION"}
    if not required.issubset(set(wb.sheetnames)):
        raise ValueError(
            f"El archivo {uploaded_file.name} no contiene las hojas mínimas esperadas: {', '.join(sorted(required))}."
        )

    ref_ws = wb["REFERENCIAS"]
    eval_ws = wb["EVALUACION"]
    mod_ws = wb["MODULO"] if "MODULO" in wb.sheetnames else None

    name = eval_ws["C2"].value or (mod_ws["H14"].value if mod_ws else None) or uploaded_file.name
    species = eval_ws["C4"].value or (mod_ws["H15"].value if mod_ws else None) or "-"
    date_value = eval_ws["C6"].value or (mod_ws["H16"].value if mod_ws else None)

    indicators_df = build_indicator_frame(ref_ws, eval_ws)
    if indicators_df["score_raw"].notna().sum() < 20:
        raise ValueError(
            f"El archivo {uploaded_file.name} no parece tener suficientes puntuaciones cargadas. "
            "Ábrelo en Excel, recalcula/guarda y vuelve a subirlo."
        )

    trunks_df = summarise_trunks(indicators_df)
    global_summary = summarise_global(indicators_df)
    strongest_trunk, weakest_trunk = strongest_and_weakest_trunk(trunks_df)

    return {
        "filename": uploaded_file.name,
        "name": str(name).strip(),
        "species": str(species).strip() if species is not None else "-",
        "date": date_value,
        "indicators": indicators_df,
        "trunks": trunks_df,
        "global": global_summary,
        "global_level": infer_level(global_summary["avg_raw"]),
        "strongest_trunk": strongest_trunk,
        "weakest_trunk": weakest_trunk,
    }


def score_candidate(candidate: dict):
    indicators_df = candidate["indicators"].copy()
    trunks_df = candidate["trunks"].copy()
    global_summary = candidate["global"]

    global_score = (
        0.60 * ratio_to_score(global_summary["vs_goal"], stretch=1.15)
        + 0.40 * ratio_to_score(global_summary["vs_max"], stretch=1.00)
    )

    trunk_goal_values = [v for v in trunks_df["vs_goal"].tolist() if v is not None and not pd.isna(v)]
    avg_trunk_goal = statistics.mean(trunk_goal_values) if trunk_goal_values else 0.0
    min_trunk_goal = min(trunk_goal_values) if trunk_goal_values else 0.0
    balance_score = (
        0.60 * ratio_to_score(avg_trunk_goal, stretch=1.10)
        + 0.40 * ratio_to_score(min_trunk_goal, stretch=1.00)
    )

    critical_df = indicators_df[indicators_df["indicator"].isin(CRITICAL_INDICATORS.keys())].copy()
    if critical_df.empty:
        critical_score = 0.0
        weak_critical_count = 999
    else:
        critical_df["priority_weight"] = critical_df["indicator"].map(CRITICAL_INDICATORS).fillna(1.0)
        critical_df["goal_score"] = critical_df["vs_goal"].apply(lambda x: ratio_to_score(x, stretch=1.15))
        critical_df["max_score"] = critical_df["vs_max"].apply(lambda x: ratio_to_score(x, stretch=1.00))
        weighted_goal = (critical_df["goal_score"] * critical_df["priority_weight"]).sum() / critical_df["priority_weight"].sum()
        weighted_max = (critical_df["max_score"] * critical_df["priority_weight"]).sum() / critical_df["priority_weight"].sum()
        weak_critical_count = int((critical_df["vs_goal"] < 0.85).fillna(False).sum())
        penalty = min(24, weak_critical_count * 8)
        critical_score = max(0.0, (0.70 * weighted_goal + 0.30 * weighted_max) - penalty)

    transfer_df = indicators_df[indicators_df["indicator"].isin(TRANSFER_INDICATORS)].copy()
    if transfer_df.empty:
        transfer_score = 0.0
    else:
        transfer_goal = transfer_df["vs_goal"].apply(lambda x: ratio_to_score(x, stretch=1.10)).mean()
        transfer_bbdd = transfer_df["vs_bbdd"].apply(lambda x: ratio_to_score(x, stretch=1.20)).mean()
        transfer_score = 0.70 * transfer_goal + 0.30 * transfer_bbdd

    above_bbdd_pct = float((indicators_df["vs_bbdd"] >= 1.0).fillna(False).mean() * 100.0)
    team_advantage_score = (
        0.70 * ratio_to_score(global_summary["vs_bbdd"], stretch=1.25)
        + 0.30 * above_bbdd_pct
    )

    senior_score = (
        RANKING_WEIGHTS["global_performance"] * global_score
        + RANKING_WEIGHTS["balance"] * balance_score
        + RANKING_WEIGHTS["critical"] * critical_score
        + RANKING_WEIGHTS["transfer"] * transfer_score
        + RANKING_WEIGHTS["team_advantage"] * team_advantage_score
    )

    min_trunk = min_trunk_goal if trunk_goal_values else 0.0
    eligibility_checks = {
        "global_vs_goal": (global_summary["vs_goal"] is not None and global_summary["vs_goal"] >= 1.00),
        "weakest_trunk": (min_trunk >= 0.90),
        "critical_gaps": (weak_critical_count <= 1),
        "transfer_floor": (transfer_score >= 55.0),
    }
    eligible_now = all(eligibility_checks.values())

    return {
        "component_global": round(global_score, 1),
        "component_balance": round(balance_score, 1),
        "component_critical": round(critical_score, 1),
        "component_transfer": round(transfer_score, 1),
        "component_team_advantage": round(team_advantage_score, 1),
        "senior_score": round(senior_score, 1),
        "eligible_now": eligible_now,
        "eligibility_checks": eligibility_checks,
        "weak_critical_count": weak_critical_count,
        "above_bbdd_pct": round(above_bbdd_pct, 1),
    }


def top_strengths(indicators_df: pd.DataFrame, top_n=3):
    tmp = indicators_df.copy()
    tmp["composite"] = (
        tmp["vs_goal"].fillna(0) * 0.60
        + tmp["vs_max"].fillna(0) * 0.25
        + tmp["vs_bbdd"].fillna(0) * 0.15
    )
    top = tmp.sort_values("composite", ascending=False).head(top_n)
    return top[["indicator", "tronco", "score_raw", "vs_goal", "vs_bbdd"]].to_dict("records")


def main_gaps(indicators_df: pd.DataFrame, top_n=3):
    tmp = indicators_df.copy()
    tmp["is_critical"] = tmp["indicator"].isin(CRITICAL_INDICATORS.keys())
    tmp["priority"] = tmp["is_critical"].astype(int)
    tmp["composite"] = (
        tmp["vs_goal"].fillna(0) * 0.70
        + tmp["vs_bbdd"].fillna(0) * 0.30
    )
    tmp = tmp.sort_values(["priority", "composite"], ascending=[False, True])
    gaps = tmp.head(top_n)
    return gaps[["indicator", "tronco", "score_raw", "vs_goal", "vs_bbdd"]].to_dict("records")


def candidate_argument(candidate: dict, score_data: dict, selected_name: str = None):
    global_summary = candidate["global"]
    strengths = top_strengths(candidate["indicators"], top_n=3)
    gaps = main_gaps(candidate["indicators"], top_n=3)

    strengths_txt = "; ".join(
        [f"{x['indicator']} ({format_pct(x['vs_goal'])} vs objetivo)" for x in strengths]
    ) or "sin fortalezas destacadas"
    gaps_txt = "; ".join(
        [f"{x['indicator']} ({format_pct(x['vs_goal'])} vs objetivo)" for x in gaps]
    ) or "sin gaps relevantes"

    if candidate["name"] == selected_name and score_data["eligible_now"]:
        return (
            f"**Elección propuesta para consultor senior.** "
            f"Obtiene el mejor Senior Score ({score_data['senior_score']}/100), "
            f"supera el objetivo global ({format_pct(global_summary['vs_goal'])}), "
            f"mantiene un perfil equilibrado y presenta fortalezas especialmente relevantes para el rol: {strengths_txt}. "
            f"Su tronco más sólido es **{candidate['strongest_trunk']}** y no muestra debilidades estructurales incompatibles con el rol."
        )

    reasons = []
    if not score_data["eligibility_checks"]["global_vs_goal"]:
        reasons.append("no supera el objetivo global")
    if not score_data["eligibility_checks"]["weakest_trunk"]:
        reasons.append(f"presenta debilidad en el tronco {candidate['weakest_trunk']}")
    if not score_data["eligibility_checks"]["critical_gaps"]:
        reasons.append("acumula carencias en indicadores críticos")
    if not score_data["eligibility_checks"]["transfer_floor"]:
        reasons.append("todavía no alcanza suficiente capacidad de transferencia/formación")

    reason_txt = ", ".join(reasons) if reasons else "queda por detrás de otros candidatos con mejor combinación global para el rol"
    return (
        f"**No propuesto como senior por ahora.** "
        f"Sus principales fortalezas son: {strengths_txt}. "
        f"No obstante, la app no lo prioriza como senior porque {reason_txt}. "
        f"Los principales focos de mejora son: {gaps_txt}."
    )


def ranking_dataframe(scored_candidates: list) -> pd.DataFrame:
    rows = []
    for idx, item in enumerate(scored_candidates, start=1):
        candidate = item["candidate"]
        score_data = item["score"]
        global_summary = candidate["global"]
        rows.append(
            {
                "Ranking": idx,
                "Nombre": candidate["name"],
                "Especie": candidate["species"],
                "Apto senior ahora": "Sí" if score_data["eligible_now"] else "No",
                "Senior Score": score_data["senior_score"],
                "Nivel global": candidate["global_level"],
                "Vs objetivo global": pct(global_summary["vs_goal"]),
                "Vs máximo global": pct(global_summary["vs_max"]),
                "Vs media BBDD": pct(global_summary["vs_bbdd"]),
                "Tronco más fuerte": candidate["strongest_trunk"],
                "Tronco más débil": candidate["weakest_trunk"],
            }
        )
    return pd.DataFrame(rows)


def detailed_export_dataframe(scored_candidates: list) -> pd.DataFrame:
    rows = []
    for item in scored_candidates:
        candidate = item["candidate"]
        score_data = item["score"]
        for _, row in candidate["indicators"].iterrows():
            rows.append(
                {
                    "Nombre": candidate["name"],
                    "Especie": candidate["species"],
                    "Tronco": row["tronco"],
                    "Indicador": row["indicator"],
                    "Score bruto": row["score_raw"],
                    "Score ponderado": row["score_weighted"],
                    "Vs objetivo": pct(row["vs_goal"]),
                    "Vs máximo": pct(row["vs_max"]),
                    "Vs media BBDD": pct(row["vs_bbdd"]),
                    "Senior Score candidato": score_data["senior_score"],
                    "Apto senior ahora": "Sí" if score_data["eligible_now"] else "No",
                }
            )
    return pd.DataFrame(rows)


def build_excel_report(scored_candidates: list):
    ranking_df = ranking_dataframe(scored_candidates)
    detail_df = detailed_export_dataframe(scored_candidates)
    method_df = CRITERIA_TABLE.copy()

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        ranking_df.to_excel(writer, sheet_name="Ranking", index=False)
        detail_df.to_excel(writer, sheet_name="Detalle", index=False)
        method_df.to_excel(writer, sheet_name="Criterios", index=False)

    bio.seek(0)
    return bio.getvalue()


# =========================================================
# APP
# =========================================================

st.title("Senior Technical Consultant Selector")
st.caption(
    "Sube entre 2 y 10 evaluaciones individuales y la app generará un ranking completo de candidatos "
    "para consultor senior, con argumentos técnicos y criterios de selección transparentes."
)

with st.expander("Criterios de selección del senior", expanded=True):
    st.markdown(CRITERIA_TEXT)
    st.table(CRITERIA_TABLE)

st.info(
    "Importante: la app funciona mejor si los ficheros se han abierto y guardado previamente en Excel "
    "con las fórmulas actualizadas."
)

uploaded_files = st.file_uploader(
    "Sube de 2 a 10 evaluaciones (.xlsm o .xlsx)",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True,
    help="Cada archivo debe corresponder a una evaluación individual completa.",
)

if not uploaded_files:
    st.stop()

if len(uploaded_files) < 2:
    st.warning("Necesitas subir al menos 2 evaluaciones para construir un ranking comparativo.")
    st.stop()

if len(uploaded_files) > 10:
    st.warning("Por ahora la app está pensada para un máximo de 10 evaluaciones por análisis.")
    st.stop()

parsed_candidates = []
errors = []

for file in uploaded_files:
    try:
        candidate = parse_candidate(file)
        score_data = score_candidate(candidate)
        parsed_candidates.append({"candidate": candidate, "score": score_data})
    except Exception as exc:
        errors.append(f"{file.name}: {exc}")

if errors:
    st.error("Algunos archivos no se han podido procesar:")
    for err in errors:
        st.write(f"- {err}")

if not parsed_candidates:
    st.stop()

parsed_candidates = sorted(
    parsed_candidates,
    key=lambda x: (
        1 if x["score"]["eligible_now"] else 0,
        x["score"]["senior_score"],
        x["candidate"]["global"]["vs_goal"] or 0,
    ),
    reverse=True,
)

selected_candidate = parsed_candidates[0]["candidate"]
selected_score = parsed_candidates[0]["score"]

st.subheader("Conclusión ejecutiva")

if selected_score["eligible_now"]:
    st.success(
        f"**Candidato propuesto para consultor senior: {selected_candidate['name']}** "
        f"(Senior Score {selected_score['senior_score']}/100)."
    )
else:
    st.warning(
        f"**No aparece un senior plenamente consolidado en este grupo.** "
        f"El mejor perfil actual es **{selected_candidate['name']}** "
        f"(Senior Score {selected_score['senior_score']}/100), pero todavía no cumple todo el estándar senior."
    )

st.markdown(candidate_argument(selected_candidate, selected_score, selected_name=selected_candidate["name"]))

st.subheader("Ranking final de candidatos")
ranking_df = ranking_dataframe(parsed_candidates)
st.dataframe(ranking_df, use_container_width=True, hide_index=True)

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Nº candidatos", len(parsed_candidates))
with col2:
    st.metric("Aptos senior ahora", sum(1 for x in parsed_candidates if x["score"]["eligible_now"]))
with col3:
    st.metric("Mejor Senior Score", f"{parsed_candidates[0]['score']['senior_score']}/100")
with col4:
    st.metric("Mejor vs objetivo global", format_pct(max((x["candidate"]["global"]["vs_goal"] or 0) for x in parsed_candidates)))
with col5:
    st.metric("Mejor vs media BBDD", format_pct(max((x["candidate"]["global"]["vs_bbdd"] or 0) for x in parsed_candidates)))

chart_df = ranking_df[["Nombre", "Senior Score"]].set_index("Nombre")
st.bar_chart(chart_df)

st.subheader("Detalle y argumentos por candidato")

for item in parsed_candidates:
    candidate = item["candidate"]
    score_data = item["score"]
    global_summary = candidate["global"]

    title = (
        f"{candidate['name']} · Senior Score {score_data['senior_score']}/100 · "
        f"{'Apto senior ahora' if score_data['eligible_now'] else 'No apto senior todavía'}"
    )
    with st.expander(title, expanded=(candidate["name"] == selected_candidate["name"])):
        left, right = st.columns([1.05, 1])

        with left:
            st.markdown(candidate_argument(candidate, score_data, selected_name=selected_candidate["name"]))

            components_df = pd.DataFrame(
                [
                    {"Componente": "Rendimiento global", "Score": score_data["component_global"]},
                    {"Componente": "Equilibrio entre troncos", "Score": score_data["component_balance"]},
                    {"Componente": "Indicadores críticos", "Score": score_data["component_critical"]},
                    {"Componente": "Transferencia y formación", "Score": score_data["component_transfer"]},
                    {"Componente": "Ventaja frente a la media", "Score": score_data["component_team_advantage"]},
                ]
            )
            st.markdown("**Desglose del ranking final**")
            st.dataframe(components_df, use_container_width=True, hide_index=True)

            checks_df = pd.DataFrame(
                [
                    {"Filtro mínimo senior": "Supera objetivo global", "Cumple": "Sí" if score_data["eligibility_checks"]["global_vs_goal"] else "No"},
                    {"Filtro mínimo senior": "Sin tronco claramente débil", "Cumple": "Sí" if score_data["eligibility_checks"]["weakest_trunk"] else "No"},
                    {"Filtro mínimo senior": "Sin carencias graves en críticos", "Cumple": "Sí" if score_data["eligibility_checks"]["critical_gaps"] else "No"},
                    {"Filtro mínimo senior": "Capacidad de transferencia suficiente", "Cumple": "Sí" if score_data["eligibility_checks"]["transfer_floor"] else "No"},
                ]
            )
            st.markdown("**Filtro mínimo senior**")
            st.dataframe(checks_df, use_container_width=True, hide_index=True)

        with right:
            metric_cols = st.columns(2)
            metric_cols[0].metric("Vs objetivo global", format_pct(global_summary["vs_goal"]))
            metric_cols[1].metric("Vs máximo global", format_pct(global_summary["vs_max"]))
            metric_cols[0].metric("Vs media BBDD", format_pct(global_summary["vs_bbdd"]))
            metric_cols[1].metric("Nivel global", candidate["global_level"])
            metric_cols[0].metric("Tronco más fuerte", candidate["strongest_trunk"])
            metric_cols[1].metric("Tronco más débil", candidate["weakest_trunk"])

            trunks_display = candidate["trunks"][["tronco", "vs_goal", "vs_max", "vs_bbdd"]].copy()
            trunks_display["vs_goal"] = trunks_display["vs_goal"].apply(format_pct)
            trunks_display["vs_max"] = trunks_display["vs_max"].apply(format_pct)
            trunks_display["vs_bbdd"] = trunks_display["vs_bbdd"].apply(format_pct)
            st.markdown("**Comparación por troncos**")
            st.dataframe(trunks_display, use_container_width=True, hide_index=True)

        indicator_display = candidate["indicators"][["tronco", "indicator", "score_raw", "weight", "vs_goal", "vs_max", "vs_bbdd"]].copy()
        indicator_display["vs_goal"] = indicator_display["vs_goal"].apply(format_pct)
        indicator_display["vs_max"] = indicator_display["vs_max"].apply(format_pct)
        indicator_display["vs_bbdd"] = indicator_display["vs_bbdd"].apply(format_pct)
        indicator_display = indicator_display.rename(
            columns={
                "tronco": "Tronco",
                "indicator": "Indicador",
                "score_raw": "Score",
                "weight": "Peso",
                "vs_goal": "Vs objetivo",
                "vs_max": "Vs máximo",
                "vs_bbdd": "Vs media BBDD",
            }
        )
        st.markdown("**Detalle por indicador**")
        st.dataframe(indicator_display, use_container_width=True, hide_index=True)

st.subheader("Descargas")
export_bytes = build_excel_report(parsed_candidates)

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "Descargar ranking y detalle en Excel",
        data=export_bytes,
        file_name="senior_consultant_ranking.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with d2:
    json_payload = {
        item["candidate"]["name"]: {
            "species": item["candidate"]["species"],
            "senior_score": item["score"]["senior_score"],
            "eligible_now": item["score"]["eligible_now"],
            "global": item["candidate"]["global"],
        }
        for item in parsed_candidates
    }
    st.download_button(
        "Descargar resumen JSON",
        data=json.dumps(json_payload, ensure_ascii=False, indent=2, default=str).encode("utf-8"),
        file_name="senior_consultant_ranking.json",
        mime="application/json",
        use_container_width=True,
    )
