
# app.py
# -*- coding: utf-8 -*-

import io
import re
import unicodedata
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st

# NOVO: imports para mapa térmico
import requests
import plotly.express as px

# -------------------------------------------------------------
# Configuração do app
# -------------------------------------------------------------
st.set_page_config(
    page_title="R.O - Impedimentos de Intervenção",
    page_icon="🚧",
    layout="wide"
)

st.title("🚧 Controle de R.O / BO - Impedimentos de Intervenção")
st.caption("By Raphael + Copilot — Streamlit")

# -------------------------------------------------------------
# Funções utilitárias
# -------------------------------------------------------------
def strip_all(s):
    if pd.isna(s):
        return s
    return str(s).replace("\n", " ").replace("\r", " ").strip()

def normalize_spaces(s):
    if pd.isna(s):
        return s
    # remove duplos espaços e espaços antes de pontuação
    s = re.sub(r"\s+", " ", str(s))
    s = re.sub(r"\s+([,.])", r"\1", s)
    return s.strip()

def normalize_text(s):
    if pd.isna(s):
        return s
    return normalize_spaces(strip_all(s))

def strip_accents(s: str) -> str:
    if s is None:
        return s
    s = str(s)
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def normalize_bo(bo):
    """Padroniza BO no formato NNN-NNNNN/AAAA a partir de entradas variadas."""
    if pd.isna(bo) or str(bo).strip() == "":
        return None
    s = str(bo)
    # mantém somente dígitos e separadores principais
    s = re.sub(r"[^0-9/\-]", "", s)
    # já está no padrão correto
    if re.fullmatch(r"\d{3}-\d{5}/\d{4}", s):
        return s
    # captura grupos de números (primeiros 3, próximos 5 e últimos 4)
    digits = re.findall(r"\d+", s)
    if len(digits) >= 3:
        a, b, c = digits[-3], digits[-2], digits[-1]
        a = a.zfill(3)[-3:]
        b = b.zfill(5)[-5:]
        c = c.zfill(4)[-4:]
        return f"{a}-{b}/{c}"
    return s if s else None

def try_parse_date_any(dval):
    """Aceita datas em dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd e variações."""
    if dval is None or (isinstance(dval, float) and pd.isna(dval)):
        return None
    # já é Timestamp/date
    if isinstance(dval, (pd.Timestamp, datetime)):
        return pd.to_datetime(dval).date()
    if isinstance(dval, date):
        return dval
    s = str(dval).strip()
    # Tenta formatos explícitos
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    # Pandas parse (tenta com dayfirst True e False)
    for dayfirst in (True, False):
        try:
            ts = pd.to_datetime(s, dayfirst=dayfirst, errors="coerce")
            if not pd.isna(ts):
                return ts.date()
        except Exception:
            pass
    return None

def google_maps_link(endereco, bairro=None, cidade=None):
    parts = [str(endereco)]
    if bairro: parts.append(str(bairro))
    if cidade: parts.append(str(cidade))
    q = ", ".join([p for p in parts if p and str(p).strip()])
    return f"https://www.google.com/maps/search/?api=1&query={q.replace(' ', '+')}"

def df_to_excel_bytes(df, filename="export.xlsx"):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtrado")
    out.seek(0)
    return out.getvalue()

def kpi_card(label, value, help_text=None):
    st.metric(label, value)
    if help_text:
        st.caption(help_text)

# ------------- GEOJSON / Choropleth helpers (NOVO) -------------
@st.cache_data(show_spinner=True)
def load_geojson_municipios(uf_code: str, allow_insecure: bool = False):
    """
    Carrega GeoJSON de municípios da UF (ex.: RJ='33', ES='32') com múltiplos fallbacks:
      1) jsDelivr (CDN do GitHub)
      2) GitHub Raw
      3) API de Malhas IBGE (estado, resolucao=5 => inclui municípios)
    Se allow_insecure=True, última tentativa ignora verificação SSL (apenas teste).
    """
    # 1) jsDelivr (CDN do repositório geodata-br)
    cdn_url = f"https://cdn.jsdelivr.net/gh/tbrugz/geodata-br/geojson/geojs-{uf_code}-mun.json"
    # 2) GitHub Raw (pode ser bloqueado por proxy)
    gh_url  = f"https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-{uf_code}-mun.json"
    # 3) API de Malhas IBGE (retorna GeoJSON; resolucao=5 inclui municípios)
    # Docs oficiais da API de malhas: servicodados.ibge.gov.br/api/docs/malhas?versao=3
    ibge_url = f"https://servicodados.ibge.gov.br/api/v3/malhas/estados/{uf_code}"
    ibge_params = {
        "resolucao": "5",
        "formato": "application/vnd.geo+json",
        "qualidade": "2"  # 1..4 (quanto maior, mais detalhado e pesado)
    }

    # Tentativas seguras (SSL verificado)
    for url in (cdn_url, gh_url):
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            return r.json()
        except Exception:
            pass  # tenta próximo

    try:
        r = requests.get(ibge_url, params=ibge_params, timeout=30)
        r.raise_for_status()
        return r.json()
    except Exception:
        pass

    # (Opcional) Última tentativa: ignorar verificação SSL (APENAS TESTE)
    if allow_insecure:
        for url in (cdn_url, gh_url):
            try:
                r = requests.get(url, timeout=30, verify=False)
                r.raise_for_status()
                st.warning("⚠️ SSL desativado para baixar o GeoJSON (use apenas para teste).")
                return r.json()
            except Exception:
                pass
        try:
            r = requests.get(ibge_url, params=ibge_params, timeout=30, verify=False)
            r.raise_for_status()
            st.warning("⚠️ SSL desativado para baixar o GeoJSON (use apenas para teste).")
            return r.json()
        except Exception:
            pass

    raise RuntimeError("Não foi possível carregar o GeoJSON de municípios desta UF pelas fontes disponíveis.")

def norm_city_name(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return strip_accents(str(s)).upper().strip()

# Correções de nomes comuns observados (ajuste aqui quando aparecerem novas divergências)
CITY_FIX = {
    "SAO JOAO DE MERIT": "SAO JOAO DE MERITI",
    "DUQUE DE CAXIA": "DUQUE DE CAXIAS",
    "RIO DE JANEIRO-": "RIO DE JANEIRO",
}

def apply_city_fix(s: str) -> str:
    base = norm_city_name(s)
    return CITY_FIX.get(base, base)

# -------------------------------------------------------------
# Upload de arquivo
# -------------------------------------------------------------
with st.sidebar:
    st.header("📎 Dados")
    st.write("Carregue sua planilha **BO_FUST.xlsx** ou equivalente.")
    file = st.file_uploader(
        "Planilha (.xlsx/.xls/.csv)",
        type=["xlsx", "xls", "csv"],
        help="Dica: sua aba principal parece se chamar 'Base_Impedimentos'."
    )
    default_sheet = st.text_input("Nome da aba (opcional)", value="Base_Impedimentos")

    st.markdown("---")
    st.header("🔧 Opções")
    enable_email = st.checkbox("Habilitar template de e-mail", value=True)
    enable_form = st.checkbox("Habilitar formulário para novo registro", value=True)
    show_quality = st.checkbox("Mostrar relatório de qualidade dos dados", value=True)

if not file:
    st.info("Carregue a planilha na barra lateral para começar.")
    st.stop()

# Leitura do arquivo
@st.cache_data(show_spinner=True)
def load_dataframe(_file, sheet_hint=None):
    fname = _file.name.lower()
    if fname.endswith((".xlsx", ".xls")):
        xls = pd.ExcelFile(_file)
        # tenta usar a sheet_hint se existir
        sheet_to_use = None
        if sheet_hint and sheet_hint in xls.sheet_names:
            sheet_to_use = sheet_hint
        else:
            # heurística: prioriza Base_Impedimentos
            for guess in ["Base_Impedimentos", "base_impedimentos", "Planilha1", "Sheet1"]:
                if guess in xls.sheet_names:
                    sheet_to_use = guess
                    break
            if sheet_to_use is None:
                # usa primeira aba
                sheet_to_use = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_to_use, engine="openpyxl")
        return df, sheet_to_use
    elif fname.endswith(".csv"):
        df = pd.read_csv(_file, sep=None, engine="python")
        return df, None
    else:
        raise ValueError("Formato não suportado. Use .xlsx, .xls ou .csv")

try:
    df_raw, used_sheet = load_dataframe(file, sheet_hint=(default_sheet or None))
except Exception as e:
    st.error(f"Erro ao ler o arquivo: {e}")
    st.stop()

st.caption(f"Arquivo carregado: **{file.name}** — aba usada: **{used_sheet or 'CSV'}**")

# -------------------------------------------------------------
# Padronização de colunas (mapeamento flexível)
# -------------------------------------------------------------
expected_cols = [
    "ID_Chamado", "Data", "Tipo_Cliente", "Nome_Cliente",
    "Endereco", "Bairro", "Cidade", "BO_Numero",
    "Tipo_Impedimento", "Projeto_Publico"
]

# cria um mapeamento por aproximação
def _map_cols(df_in: pd.DataFrame) -> dict:
    cmap = {}
    for col in df_in.columns:
        c = str(col).strip()
        cl = strip_accents(c.lower())
        if cl.startswith("id_cha") or cl.startswith("idch"): cmap[col] = "ID_Chamado"
        elif cl == "data" or cl.startswith("data"): cmap[col] = "Data"
        elif ("tipo" in cl and "cliente" in cl) or cl == "tipo_cliente": cmap[col] = "Tipo_Cliente"
        elif ("nome" in cl and "cliente" in cl) or cl == "nome_cliente": cmap[col] = "Nome_Cliente"
        elif "ender" in cl: cmap[col] = "Endereco"
        elif "bairro" in cl: cmap[col] = "Bairro"
        elif "cidade" in cl: cmap[col] = "Cidade"
        elif cl.startswith("bo") or "bo_num" in cl: cmap[col] = "BO_Numero"
        elif "imped" in cl: cmap[col] = "Tipo_Impedimento"
        elif "projeto" in cl: cmap[col] = "Projeto_Publico"
    return cmap

colmap = _map_cols(df_raw)
df = df_raw.rename(columns=colmap).copy()
missing = [c for c in expected_cols if c not in df.columns]
if missing:
    st.warning(
        "Algumas colunas obrigatórias não foram encontradas: " + ", ".join(missing) +
        "\n\nVou tentar continuar com o que existir, mas certos recursos podem ficar limitados."
    )

# garante todas as colunas esperadas (criando vazias se não existirem)
for c in expected_cols:
    if c not in df.columns:
        df[c] = None

# -------------------------------------------------------------
# Limpeza e enriquecimento
# -------------------------------------------------------------
text_cols = [
    "Tipo_Cliente", "Nome_Cliente", "Endereco", "Bairro", "Cidade",
    "Tipo_Impedimento", "Projeto_Publico", "BO_Numero"
]
for c in text_cols:
    df[c] = df[c].apply(normalize_text)

# Datas
df["Data_Parsed"] = df["Data"].apply(try_parse_date_any)

# BO normalizado
df["BO_Numero_Normalizado"] = df["BO_Numero"].apply(normalize_bo)

# Projeto Público normalizado (SIM/NÃO)
pp = df["Projeto_Publico"].fillna("").astype(str).str.upper()
pp = pp.apply(lambda x: strip_accents(x))
pp = pp.replace({
    "SIM": "SIM",
    "NAO": "NÃO",
    "NAO ": "NÃO",
    "N/A": "NÃO",
    "": None
})
df["Projeto_Publico"] = pp

# Link Google Maps
df["Google_Maps"] = df.apply(
    lambda r: google_maps_link(r.get("Endereco"), r.get("Bairro"), r.get("Cidade")), axis=1
)

# Duplicidades
df["Duplicado_ID"] = df["ID_Chamado"].astype(str).duplicated(keep=False)
# trata BO nulo como não duplicado
bo_tmp = df["BO_Numero_Normalizado"].fillna("__NA__").astype(str)
df["Duplicado_BO"] = bo_tmp.duplicated(keep=False) & (bo_tmp != "__NA__")

# -------------------------------------------------------------
# Filtros
# -------------------------------------------------------------
st.subheader("🔎 Filtros")

col1, col2, col3, col4 = st.columns(4)
with col1:
    min_date = df["Data_Parsed"].min()
    max_date = df["Data_Parsed"].max()
    date_range = st.date_input(
        "Período (Data do R.O)",
        value=(min_date, max_date) if (min_date and max_date) else None
    )
with col2:
    tipos = ["(Todos)"] + sorted([x for x in df["Tipo_Cliente"].dropna().unique()])
    tipo_sel = st.selectbox("Tipo de Cliente", tipos)
with col3:
    bairros = ["(Todos)"] + sorted([x for x in df["Bairro"].dropna().unique()])
    bairro_sel = st.selectbox("Bairro", bairros)
with col4:
    cidades = ["(Todos)"] + sorted([x for x in df["Cidade"].dropna().unique()])
    cidade_sel = st.selectbox("Cidade", cidades)

col5, col6, col7 = st.columns(3)
with col5:
    impedimentos = ["(Todos)"] + sorted([x for x in df["Tipo_Impedimento"].dropna().unique()])
    imp_sel = st.selectbox("Tipo de Impedimento", impedimentos)
with col6:
    proj_opts = ["(Todos)", "SIM", "NÃO"]
    proj_sel = st.selectbox("Projeto Público", proj_opts)
with col7:
    quick_search = st.text_input("Busca rápida (ID, BO, Nome, Endereço)", value="").strip()

# aplica filtros
fdf = df.copy()
if isinstance(date_range, tuple) and len(date_range) == 2 and all(date_range):
    d0, d1 = date_range
    fdf = fdf[(fdf["Data_Parsed"] >= d0) & (fdf["Data_Parsed"] <= d1)]
if tipo_sel != "(Todos)":
    fdf = fdf[fdf["Tipo_Cliente"] == tipo_sel]
if bairro_sel != "(Todos)":
    fdf = fdf[fdf["Bairro"] == bairro_sel]
if cidade_sel != "(Todos)":
    fdf = fdf[fdf["Cidade"] == cidade_sel]
if imp_sel != "(Todos)":
    fdf = fdf[fdf["Tipo_Impedimento"] == imp_sel]
if proj_sel != "(Todos)":
    fdf = fdf[fdf["Projeto_Publico"].fillna("NÃO") == proj_sel]

if quick_search:
    qs = quick_search.lower()
    mask = (
        fdf["ID_Chamado"].astype(str).str.contains(qs, case=False, na=False) |
        fdf["BO_Numero"].astype(str).str.contains(qs, case=False, na=False) |
        fdf["BO_Numero_Normalizado"].astype(str).str.contains(qs, case=False, na=False) |
        fdf["Nome_Cliente"].astype(str).str.lower().str.contains(qs, na=False) |
        fdf["Endereco"].astype(str).str.lower().str.contains(qs, na=False)
    )
    fdf = fdf[mask]

# -------------------------------------------------------------
# KPIs (sem gráficos, apenas métricas)
# -------------------------------------------------------------
st.subheader("📊 Indicadores")
c1, c2, c3, c4 = st.columns(4)
with c1:
    kpi_card("Total de R.O no filtro", f"{len(fdf):,}".replace(",", "."))
with c2:
    top_imp = fdf["Tipo_Impedimento"].value_counts().head(1)
    kpi_card("Impedimento mais comum", f"{top_imp.index[0]} ({int(top_imp.iloc[0])})" if not top_imp.empty else "-")
with c3:
    sim = (fdf["Projeto_Publico"].fillna("NÃO") == "SIM").sum()
    kpi_card("Projeto Público (SIM)", f"{sim}")
with c4:
    uniques_bo = fdf["BO_Numero_Normalizado"].nunique()
    kpi_card("BO distintos", f"{uniques_bo}")

# -------------------------------------------------------------
# Tabela principal (Maps como botão na 1ª coluna)
# -------------------------------------------------------------
st.subheader("📋 Registros (com botão de Maps)")

# Colunas com Google Maps no início
cols_show = [
    "Google_Maps",  # <- Maps primeiro
    "ID_Chamado","Data","Tipo_Cliente","Nome_Cliente","Endereco","Bairro","Cidade",
    "BO_Numero","BO_Numero_Normalizado","Tipo_Impedimento","Projeto_Publico",
    "Duplicado_ID","Duplicado_BO"
]
# garante colunas presentes
cols_show = [c for c in cols_show if c in fdf.columns]

# Destaque de duplicidade (Styler)
from pandas.io.formats.style import Styler
def highlight_dups_col(series):
    return ['background-color: #ffd6d6' if bool(v) else '' for v in series]

styled = fdf[cols_show].style
if "Duplicado_ID" in fdf.columns:
    styled = styled.apply(highlight_dups_col, subset=["Duplicado_ID"])  # type: ignore
if "Duplicado_BO" in fdf.columns:
    styled = styled.apply(highlight_dups_col, subset=["Duplicado_BO"])  # type: ignore

# Renderiza com coluna de link como botão
st.dataframe(
    styled,
    use_container_width=True,
    height=480,
    column_config={
        "Google_Maps": st.column_config.LinkColumn(
            label="🗺️ Maps",
            help="Abrir endereço no Google Maps",
            display_text="Abrir"
        )
    }
)

# -------------------------------------------------------------
# Mapa térmico por cidade (choropleth) - NOVO
# -------------------------------------------------------------
st.subheader("🗺️ Mapa térmico por cidade (choropleth)")

# Apenas RJ e ES conforme pedido
uf_label_to_code = {"RJ": "33", "ES": "32"}  # GeoJSON por UF no repositório geodata-br (base IBGE) / API IBGE
col_map_1, col_map_2 = st.columns([1, 3])
with col_map_1:
    uf_escolhida = st.selectbox("UF do mapa", list(uf_label_to_code.keys()), index=0)
with col_map_2:
    st.caption("Cada município é colorido pela quantidade de R.O no *filtro atual*. Use os filtros acima para mudar o mapa.")

# Agregado por cidade (normalizado) do filtro atual
city_counts = (
    fdf.assign(Cidade_norm=fdf["Cidade"].apply(apply_city_fix))
       .groupby("Cidade_norm")
       .size()
       .reset_index(name="Qtde")
)

# Carrega GeoJSON da UF escolhida (com fallbacks)
gj = None
try:
    uf_code = uf_label_to_code[uf_escolhida]
    gj = load_geojson_municipios(uf_code)  # allow_insecure=False (padrão)
except Exception as e:
    st.error(f"Falha ao carregar GeoJSON da UF {uf_escolhida}: {e}")

if gj:
    # Extrai (id, nome) do geojson
    features = gj.get("features", [])
    muni_df = pd.DataFrame([{
        "id": (ft.get("id") or ft.get("properties", {}).get("id") or ft.get("properties", {}).get("code")),
        "NM_MUN": (ft.get("properties", {}).get("name") or
                   ft.get("properties", {}).get("NM_MUNICIPIO") or
                   ft.get("properties", {}).get("name_muni"))
    } for ft in features])

    # Normaliza nome p/ join
    muni_df["Cidade_norm"] = muni_df["NM_MUN"].apply(norm_city_name)

    # Join: municípios da malha X contagem do filtro
    joined = muni_df.merge(city_counts, on="Cidade_norm", how="left").fillna({"Qtde": 0})

    # Diagnóstico de cidades não casadas
    nao_casaram = sorted(set(city_counts["Cidade_norm"]) - set(muni_df["Cidade_norm"]))
    if nao_casaram:
        with st.expander("⚠️ Cidades do filtro que não casaram com a malha (adicione no CITY_FIX se necessário)"):
            st.write(nao_casaram)

    # Choropleth
    max_q = int(joined["Qtde"].max()) if not joined.empty else 1
    fig = px.choropleth(
        joined,
        geojson=gj,
        locations="id",              # chave do município no GeoJSON
        color="Qtde",
        featureidkey="id",           # no geodata-br o ID do município está em 'id'
        color_continuous_scale="OrRd",
        range_color=(0, max_q if max_q > 0 else 1),
        hover_data={
            "NM_MUN": True,
            "Qtde": True,
            "id": False,
            "Cidade_norm": False
        },
        labels={"Qtde": "Chamados"}
    )
    fig.update_geos(fitbounds="locations", visible=False)
    fig.update_layout(
        margin=dict(l=0, r=0, t=10, b=0),
        coloraxis_colorbar=dict(title="R.O", thickness=12, len=0.7)
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Mapa indisponível no momento. Tente novamente.")

# -------------------------------------------------------------
# Exportação
# -------------------------------------------------------------
st.markdown("### ⬇️ Exportar resultado filtrado")
excel_bytes = df_to_excel_bytes(fdf[cols_show])
st.download_button(
    label="Baixar Excel (filtrado)",
    data=excel_bytes,
    file_name=f"RO_filtrado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# -------------------------------------------------------------
# Template de e-mail
# -------------------------------------------------------------
if enable_email:
    st.subheader("📧 Template de e-mail (copiar e colar)")
    options = fdf["ID_Chamado"].astype(str).tolist()
    if options:
        sel_id = st.selectbox("Selecione o ID_Chamado para compor o e-mail:", options)
        row = fdf[fdf["ID_Chamado"].astype(str) == sel_id].iloc[0]

        assunto = f"[R.O] Impedimento de Intervenção - ID {row['ID_Chamado']} - {row['Tipo_Impedimento']}"
        corpo = f"""Prezados,

Conforme verificado no local, **não é possível realizar intervenção** por motivo de **{row['Tipo_Impedimento']}**.

**Dados do chamado**
- ID do Chamado: {row['ID_Chamado']}
- Data: {row['Data']}
- Tipo de Cliente: {row['Tipo_Cliente']}
- Nome do Cliente: {row['Nome_Cliente']}
- Endereço: {row['Endereco']} - {row['Bairro']} - {row['Cidade']}
- BO (informado): {row['BO_Numero']}
- BO (normalizado): {row['BO_Numero_Normalizado']}
- Projeto Público: {row['Projeto_Publico']}
- Localização (Maps): {row['Google_Maps']}

Fico à disposição para quaisquer esclarecimentos.

Atenciosamente,
Raphael
"""
        st.text_input("Assunto", value=assunto)
        st.text_area("Corpo do e-mail", value=corpo, height=260)
    else:
        st.info("Nenhum registro no filtro para gerar e-mail.")

# -------------------------------------------------------------
# Formulário para novo registro (não persiste no arquivo original)
# -------------------------------------------------------------
if enable_form:
    st.subheader("➕ Incluir novo registro")
    with st.form("novo_registro"):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_id = st.text_input("ID_Chamado *")
            new_data = st.date_input("Data *", value=datetime.today())
            new_tipo_cli = st.text_input("Tipo_Cliente *", value="")
        with c2:
            new_nome = st.text_input("Nome_Cliente *", value="")
            new_bairro = st.text_input("Bairro *", value="")
            new_cidade = st.text_input("Cidade *", value="Rio de Janeiro")
        with c3:
            new_end = st.text_input("Endereco *", value="")
            new_bo = st.text_input("BO_Numero", value="")
            new_imp = st.text_input("Tipo_Impedimento *", value="Area de Risco")

        new_proj = st.selectbox("Projeto_Publico", ["SIM", "NÃO"], index=0)

        submitted = st.form_submit_button("Adicionar à tabela")
        if submitted:
            errors = []
            required = {
                "ID_Chamado": new_id,
                "Data": new_data,
                "Tipo_Cliente": new_tipo_cli,
                "Nome_Cliente": new_nome,
                "Endereco": new_end,
                "Bairro": new_bairro,
                "Cidade": new_cidade,
                "Tipo_Impedimento": new_imp,
            }
            for k, v in required.items():
                if not str(v).strip():
                    errors.append(f"Campo obrigatório: {k}")

            if str(new_id).strip() in df["ID_Chamado"].astype(str).tolist():
                errors.append("ID_Chamado já existe na planilha.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                new_row = {
                    "ID_Chamado": new_id,
                    "Data": new_data.strftime("%d/%m/%Y"),
                    "Tipo_Cliente": new_tipo_cli,
                    "Nome_Cliente": new_nome,
                    "Endereco": new_end,
                    "Bairro": new_bairro,
                    "Cidade": new_cidade,
                    "BO_Numero": new_bo,
                    "Tipo_Impedimento": new_imp,
                    "Projeto_Publico": new_proj,
                    "BO_Numero_Normalizado": normalize_bo(new_bo),
                    "Data_Parsed": new_data,
                    "Google_Maps": google_maps_link(new_end, new_bairro, new_cidade),
                    "Duplicado_ID": False,
                    "Duplicado_BO": False,
                }
                st.session_state.setdefault("new_rows", [])
                st.session_state["new_rows"].append(new_row)
                st.success("Registro adicionado na sessão. Para persistir, exporte o Excel filtrado e substitua sua planilha.")

    # Exibe adicionados na sessão
    if "new_rows" in st.session_state and st.session_state["new_rows"]:
        st.info("Registros adicionados nesta sessão (não alteram o arquivo original). Exporte para salvar.")
        st.dataframe(pd.DataFrame(st.session_state["new_rows"]), use_container_width=True)

# -------------------------------------------------------------
# Qualidade dos dados
# -------------------------------------------------------------
if show_quality:
    st.subheader("🩺 Qualidade dos dados")
    q1, q2 = st.columns(2)
    with q1:
        st.markdown("**Valores ausentes por coluna**")
        st.dataframe(df[expected_cols].isna().sum().to_frame("Nulos"), use_container_width=True)
    with q2:
        st.markdown("**BO em formato normalizado inválido (não 'NNN-NNNNN/AAAA')**")
        bad_bo = ~df["BO_Numero_Normalizado"].fillna("").astype(str).str.match(r"^\d{3}-\d{5}/\d{4}$")
        st.write(df.loc[bad_bo, ["ID_Chamado","BO_Numero","BO_Numero_Normalizado"]])

    st.markdown("**Chamados com duplicidade**")
    dups = df[(df["Duplicado_ID"]) | (df["Duplicado_BO"])][["ID_Chamado","BO_Numero","BO_Numero_Normalizado","Duplicado_ID","Duplicado_BO"]]
    if dups.empty:
        st.success("Nenhuma duplicidade encontrada.")
    else:
        st.dataframe(dups, use_container_width=True)

st.caption("© 2026 — Ferramenta interna para gestão de R.O e impedimentos.")