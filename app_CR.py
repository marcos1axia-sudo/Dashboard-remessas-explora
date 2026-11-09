# -*- coding: utf-8 -*-
"""
Dashboard de Tempo de Processamento de Remessas
================================================

Aplicativo Streamlit para análise interativa do tempo de processamento
das remessas em cada etapa do fluxo:

    Envio → Notificação → Pagamento → Início da Preparação → Finalização

Requisitos:
    pip install streamlit pandas plotly openpyxl xlsxwriter

Execução:
    streamlit run app_CR.py

IMPORTANTE: o dashboard não lê nenhum arquivo fixo do computador. Ao abrir o
app, use o campo de upload na barra lateral para enviar o arquivo
`Controle_de_Remessas_Explora.xlsx` (ou similar, mesmo layout de colunas).
"""

import io
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIGURAÇÃO GERAL DA PÁGINA
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard de Tempo de Processamento",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Mapeamento das colunas técnicas -> nomes amigáveis (não altera o df original)
# As etapas "Envio → Recebimento" e "Recebimento → Notificação" foram unificadas
# em uma única etapa: "Envio → Notificação".
ETAPAS = {
    "Dias envio_notificacao": "Envio → Notificação",
    "Dias notificacao_pagamento": "Notificação → Pagamento",
    "Dias pagamento_inicio_preparacao": "Pagamento → Início da Preparação",
    "Dias inicio_preparacao_finalizacao": "Início da Preparação → Finalização",
}
COLS_DIFF = list(ETAPAS.keys())

COLUNAS_DATA = [
    "DATA_ENVIO",
    "DATA_RECBTO",
    "DATA_REPORTE",
    "DATA_REG",
    "DATA_NTI_PGTO",
    "DATA_PGTO_PREP",
    "DATA_INIC",
    "DATA_PGTO_ENSAIO",
]

# Cores fixas por etapa (usadas em todos os gráficos p/ consistência visual)
CORES_ETAPAS = {
    "Envio → Notificação": "#4C78A8",
    "Notificação → Pagamento": "#54A24B",
    "Pagamento → Início da Preparação": "#E45756",
    "Início da Preparação → Finalização": "#72B7B2",
}

# Cores para a comparação por categoria RUSH
CORES_RUSH = {
    "Sim": "#E45756",
    "Não": "#4C78A8",
}

QTD_POR_PAGINA = 10  # quantidade de remessas exibidas por "página" no gráfico

# Remessas que devem ser ignoradas no cálculo da média por etapa (mas continuam
# aparecendo normalmente nos demais gráficos/tabelas). Comparação é feita de
# forma normalizada (aceita "37", "037", "37.0", etc. como a mesma remessa).
REMESSAS_EXCLUIDAS_MEDIA = {"037"}


def _normaliza_remessa(valor) -> str:
    """Normaliza o identificador da remessa para comparação (ex.: 37, "37", "37.0" -> "037")."""
    texto = str(valor).strip().split(".")[0]
    try:
        return f"{int(texto):03d}"
    except ValueError:
        return texto


def formatar_datas_para_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna uma cópia do df com todas as colunas datetime exibidas sem a hora
    (apenas dd/mm/aaaa) — usado só para exibição em tabelas, não altera os
    dados usados nos cálculos."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.date
    return df


# ==============================================================================
# 1. CARGA E PREPARAÇÃO DOS DADOS
# ==============================================================================
# Colunas sem as quais o dashboard não consegue calcular as etapas
COLUNAS_OBRIGATORIAS = [
    "REMESSA",
    "DATA_ENVIO",
    "DATA_NTI_PGTO",
    "DATA_PGTO_PREP",
    "DATA_INIC",
    "DATA_REPORTE",
]


def ler_planilha_remessas(arquivo) -> pd.DataFrame:
    """
    Lê a primeira aba do Excel localizando automaticamente a linha de
    cabeçalho (a primeira linha, entre as 20 iniciais, que contém "REMESSA"
    e "DATA_ENVIO"). Assim o app funciona mesmo que a planilha ganhe ou perca
    linhas de legenda acima da tabela.
    """
    previa = pd.read_excel(arquivo, header=None, nrows=20)
    linha_cabecalho = None
    for i, linha in previa.iterrows():
        valores = {str(v).strip() for v in linha.dropna()}
        if {"REMESSA", "DATA_ENVIO"} <= valores:
            linha_cabecalho = i
            break
    if linha_cabecalho is None:
        raise ValueError(
            'não encontrei a linha de cabeçalho (com "REMESSA" e "DATA_ENVIO") '
            "nas 20 primeiras linhas da primeira aba."
        )

    if hasattr(arquivo, "seek"):
        arquivo.seek(0)
    df = pd.read_excel(arquivo, header=linha_cabecalho)
    df.columns = [str(c).strip() for c in df.columns]
    # Descarta linhas totalmente vazias e sem número de remessa
    df = df.dropna(how="all")
    df = df[df["REMESSA"].notna()].reset_index(drop=True)
    return df


def carregar_df() -> pd.DataFrame:
    """
    Ponto de entrada dos dados reais.

    O dashboard NUNCA gera dados fictícios: ele só trabalha com um arquivo
    Excel enviado pelo próprio usuário na tela (upload). A linha de cabeçalho
    é localizada automaticamente (ver `ler_planilha_remessas`).
    """
    st.sidebar.header("📁 Arquivo de Dados")
    arquivo_enviado = st.sidebar.file_uploader(
        "Envie o arquivo Controle_de_Remessas_Explora.xlsx",
        type=["xlsx", "xls"],
        key="upload_arquivo_remessas",
    )

    if arquivo_enviado is not None:
        try:
            df_raw = ler_planilha_remessas(arquivo_enviado)
        except Exception as e:
            st.error(f"Não foi possível ler o arquivo enviado: {e}")
            st.stop()

        faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df_raw.columns]
        if faltantes:
            st.error(
                "O arquivo enviado não tem a(s) coluna(s) obrigatória(s): "
                f"{', '.join(faltantes)}. Confira se é a planilha de controle de remessas."
            )
            st.stop()
        st.session_state["df_original"] = df_raw
        st.session_state["nome_arquivo"] = arquivo_enviado.name

    if "df_original" not in st.session_state:
        st.info(
            "⬅️ Envie o arquivo `Controle_de_Remessas_Explora.xlsx` na barra lateral "
            "para carregar o dashboard."
        )
        st.stop()

    st.sidebar.caption(f"✅ Arquivo carregado: {st.session_state.get('nome_arquivo', 'arquivo enviado')}")
    return st.session_state["df_original"].copy()


def preparar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """Converte as colunas de data para datetime (sem alterar nomes de colunas)."""
    df = df.copy()
    for coluna in COLUNAS_DATA:
        if coluna in df.columns:
            # format="mixed" permite que cada valor da coluna seja interpretado
            # com seu próprio formato de data (evita o erro "Cannot infer format,
            # please specify a format" quando a coluna tem formatos inconsistentes,
            # ex.: algumas células com data e outras com texto/serial do Excel).
            df[coluna] = pd.to_datetime(
                df[coluna], errors="coerce", dayfirst=True, format="mixed"
            )

    # Normaliza a coluna RUSH para os valores canônicos "Sim" / "Não", quando existir
    if "RUSH" in df.columns:
        df["RUSH"] = (
            df["RUSH"]
            .astype(str)
            .str.strip()
            .str.capitalize()
            .replace({"Nan": np.nan, "None": np.nan, "": np.nan})
        )
        df.loc[~df["RUSH"].isin(["Sim", "Não"]), "RUSH"] = np.nan

    return df


def calcular_diferencas(df: pd.DataFrame) -> pd.DataFrame:
    """Cria as quatro colunas de diferença de dias entre etapas e o total."""
    df = df.copy()

    df["Dias envio_notificacao"] = abs((df["DATA_NTI_PGTO"] - df["DATA_ENVIO"]).dt.days)
    df["Dias notificacao_pagamento"] = abs((df["DATA_PGTO_PREP"] - df["DATA_NTI_PGTO"]).dt.days)
    df["Dias pagamento_inicio_preparacao"] = abs((df["DATA_INIC"] - df["DATA_PGTO_PREP"]).dt.days)
    df["Dias inicio_preparacao_finalizacao"] = abs((df["DATA_REPORTE"] - df["DATA_INIC"]).dt.days)

    # Dias sempre como inteiros (Int64 aceita vazios/NaN sem virar float)
    for col in COLS_DIFF:
        df[col] = df[col].round().astype("Int64")

    # Marca, por etapa, se a diferença é negativa (inconsistência de datas)
    for col in COLS_DIFF:
        df[f"{col}__inconsistente"] = (df[col] < 0).fillna(False).astype(bool)

    # Marca, por etapa, se está incompleta (alguma das datas está ausente)
    df["etapa1_incompleta"] = df["DATA_ENVIO"].isna() | df["DATA_NTI_PGTO"].isna()
    df["etapa2_incompleta"] = df["DATA_NTI_PGTO"].isna() | df["DATA_PGTO_PREP"].isna()
    df["etapa3_incompleta"] = df["DATA_PGTO_PREP"].isna() | df["DATA_INIC"].isna()
    df["etapa4_incompleta"] = df["DATA_INIC"].isna() | df["DATA_REPORTE"].isna()

    df["Processo_incompleto"] = df[
        ["etapa1_incompleta", "etapa2_incompleta", "etapa3_incompleta", "etapa4_incompleta"]
    ].any(axis=1)

    # Inconsistência = qualquer etapa com diferença negativa
    cols_inconsist = [f"{c}__inconsistente" for c in COLS_DIFF]
    df["Possui_inconsistencia"] = df[cols_inconsist].any(axis=1)

    # Dias_total: soma das quatro diferenças, tratando NaN como 0 (mas sem
    # considerar automaticamente negativos como válidos — ver validação acima)
    diffs_validas = df[COLS_DIFF].copy()
    # Para o total, negativos não são zerados aqui (são sinalizados à parte),
    # porém não contaminam a média por etapa (tratada em criar_grafico_media_etapas)
    df["Dias_total"] = diffs_validas.fillna(0).sum(axis=1).astype(int)

    # Etapa que mais demora, por remessa (ignora NaN)
    def etapa_maior_tempo(row):
        valores = row[COLS_DIFF]
        valores_validos = valores.dropna()
        if valores_validos.empty:
            return "Sem dados"
        return ETAPAS[valores_validos.idxmax()]

    df["Etapa_maior_tempo"] = df.apply(etapa_maior_tempo, axis=1)

    return df


def classificar_remessas(df: pd.DataFrame, limite_atencao: int, limite_critico: int) -> pd.DataFrame:
    """Classifica cada remessa em Normal / Atenção / Crítica conforme limites do usuário."""
    df = df.copy()

    def status(dias_total):
        if pd.isna(dias_total):
            return "⚪ Sem dados"
        if dias_total >= limite_critico:
            return "🔴 Crítica"
        elif dias_total >= limite_atencao:
            return "🟡 Atenção"
        else:
            return "🟢 Normal"

    df["Classificação"] = df["Dias_total"].apply(status)
    return df


# ==============================================================================
# 2. FILTROS
# ==============================================================================
def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    """Cria a barra lateral com filtros interativos e retorna o df filtrado."""
    st.sidebar.header("🔎 Filtros")

    df_filtrado = df.copy()

    filtros_categoricos = ["SITUACAO", "PROJETO", "RESPONSÁVEL", "TIPO", "DESTINO", "REMESSA"]

    for col in filtros_categoricos:
        if col in df.columns:
            opcoes = sorted(df[col].dropna().unique().tolist())
            selecionados = st.sidebar.multiselect(
                f"{col.title()}",
                options=["Selecionar todos"] + opcoes,
                default=["Selecionar todos"],
                key=f"filtro_{col}",
            )
            if "Selecionar todos" not in selecionados and len(selecionados) > 0:
                df_filtrado = df_filtrado[df_filtrado[col].isin(selecionados)]

    st.sidebar.markdown("---")

    # Período de DATA_ENVIO
    if "DATA_ENVIO" in df.columns and df["DATA_ENVIO"].notna().any():
        min_env = df["DATA_ENVIO"].min()
        max_env = df["DATA_ENVIO"].max()
        periodo_envio = st.sidebar.date_input(
            "Período de Envio",
            value=(min_env.date(), max_env.date()),
            min_value=min_env.date(),
            max_value=max_env.date(),
            key="periodo_envio",
        )
        if isinstance(periodo_envio, tuple) and len(periodo_envio) == 2:
            ini, fim = periodo_envio
            mask = df_filtrado["DATA_ENVIO"].isna() | (
                (df_filtrado["DATA_ENVIO"].dt.date >= ini) & (df_filtrado["DATA_ENVIO"].dt.date <= fim)
            )
            df_filtrado = df_filtrado[mask]

    # Período de DATA_REG
    if "DATA_REG" in df.columns and df["DATA_REG"].notna().any():
        min_reg = df["DATA_REG"].min()
        max_reg = df["DATA_REG"].max()
        periodo_reg = st.sidebar.date_input(
            "Período de Registro (Recebimento)",
            value=(min_reg.date(), max_reg.date()),
            min_value=min_reg.date(),
            max_value=max_reg.date(),
            key="periodo_reg",
        )
        if isinstance(periodo_reg, tuple) and len(periodo_reg) == 2:
            ini, fim = periodo_reg
            mask = df_filtrado["DATA_REG"].isna() | (
                (df_filtrado["DATA_REG"].dt.date >= ini) & (df_filtrado["DATA_REG"].dt.date <= fim)
            )
            df_filtrado = df_filtrado[mask]

    st.sidebar.markdown("---")
    st.sidebar.header("⚠️ Limites de Classificação")
    limite_atencao = st.sidebar.number_input(
        "Limite de atenção (dias)", min_value=0, value=30, step=1, key="limite_atencao"
    )
    limite_critico = st.sidebar.number_input(
        "Limite crítico (dias)", min_value=0, value=60, step=1, key="limite_critico"
    )
    if limite_critico < limite_atencao:
        st.sidebar.warning("O limite crítico é menor que o limite de atenção. Ajuste os valores.")

    return df_filtrado, limite_atencao, limite_critico


# ==============================================================================
# 3. KPIS
# ==============================================================================
def criar_kpis(df: pd.DataFrame) -> None:
    """Renderiza os cards de indicadores principais no topo do dashboard.

    A quantidade de remessas considera todas as remessas filtradas, mas as
    médias/mediana consideram apenas remessas finalizadas (DATA_REPORTE preenchida).
    """
    qtd_remessas = len(df)

    # Condicional: apenas remessas com DATA_REPORTE preenchida entram nas médias
    if "DATA_REPORTE" in df.columns:
        df_final = df[df["DATA_REPORTE"].notna()]
    else:
        df_final = df

    media_dias = df_final["Dias_total"].mean()
    mediana_dias = df_final["Dias_total"].median()

    if "REMESSA" in df_final.columns and REMESSAS_EXCLUIDAS_MEDIA:
        df_sem_037 = df_final[
            ~df_final["REMESSA"].apply(_normaliza_remessa).isin(REMESSAS_EXCLUIDAS_MEDIA)
        ]
    else:
        df_sem_037 = df_final
    media_dias_sem_037 = df_sem_037["Dias_total"].mean()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Remessas", f"{qtd_remessas}")
    c2.metric("📊 Média (dias)", f"{media_dias:.0f}" if pd.notna(media_dias) else "—")
    c3.metric("📈 Mediana (dias)", f"{mediana_dias:.0f}" if pd.notna(mediana_dias) else "—")
    c4.metric(
        f"📊 Média sem remessa {', '.join(sorted(REMESSAS_EXCLUIDAS_MEDIA))} (dias)",
        f"{media_dias_sem_037:.0f}" if pd.notna(media_dias_sem_037) else "—",
    )


# ==============================================================================
# 4. GRÁFICO 1 — MÉDIA DE DIAS POR ETAPA (com recorte por RUSH)
# ==============================================================================
def _media_por_etapa(df: pd.DataFrame) -> list:
    """Calcula, para um recorte de df já filtrado, a média/qtd por etapa."""
    linhas = []
    for col, nome in ETAPAS.items():
        serie_valida = df[col].dropna()
        serie_valida_pos = serie_valida[serie_valida >= 0]
        media = serie_valida_pos.mean() if not serie_valida_pos.empty else np.nan
        linhas.append(
            {
                "Etapa": nome,
                "Média": round(media) if pd.notna(media) else 0,
                "Qtd": int(serie_valida_pos.count()),
                "Rotulo": f"{media:.0f}" if pd.notna(media) else "Sem dados",
            }
        )
    return linhas


def criar_grafico_media_etapas(df: pd.DataFrame, rush_opcao: str) -> go.Figure:
    """
    Gráfico de barras com a média de dias por etapa (ignora NaN).

    Considera apenas remessas finalizadas (DATA_REPORTE preenchida).

    `rush_opcao` controla o recorte:
        - "Todas (comparar Sim x Não)": barras agrupadas por etapa, uma para
          cada categoria de RUSH, lado a lado.
        - "Sim" / "Não": considera apenas remessas daquela categoria.
    """
    if "DATA_REPORTE" in df.columns:
        df = df[df["DATA_REPORTE"].notna()]

    if "REMESSA" in df.columns and REMESSAS_EXCLUIDAS_MEDIA:
        df = df[~df["REMESSA"].apply(_normaliza_remessa).isin(REMESSAS_EXCLUIDAS_MEDIA)]

    tem_rush = "RUSH" in df.columns

    if not tem_rush or rush_opcao == "Todas (comparar Sim x Não)" and not tem_rush:
        # Sem coluna RUSH disponível: comportamento original (sem recorte)
        df_media = pd.DataFrame(_media_por_etapa(df))
        fig = px.bar(
            df_media, x="Etapa", y="Média", color="Etapa",
            color_discrete_map=CORES_ETAPAS, text="Rotulo", custom_data=["Qtd"],
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Média: %{text} dias<br>Remessas consideradas: %{customdata[0]}<extra></extra>",
        )
        fig.update_layout(showlegend=False, yaxis_title="Média de dias", yaxis_tickformat="d",
                           xaxis_title="", margin=dict(t=30, b=10), height=420)
        return fig

    if rush_opcao == "Todas (comparar Sim x Não)":
        partes = []
        for categoria in ["Sim", "Não"]:
            subset = df[df["RUSH"] == categoria]
            linhas = _media_por_etapa(subset)
            for linha in linhas:
                linha["RUSH"] = categoria
            partes.extend(linhas)
        df_media = pd.DataFrame(partes)

        fig = px.bar(
            df_media, x="Etapa", y="Média", color="RUSH", barmode="group",
            color_discrete_map=CORES_RUSH, text="Rotulo", custom_data=["Qtd"],
        )
        fig.update_traces(
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>RUSH: %{fullData.name}<br>Média: %{text} dias<br>Remessas consideradas: %{customdata[0]}<extra></extra>",
        )
        fig.update_layout(
            legend_title="RUSH", yaxis_title="Média de dias", yaxis_tickformat="d",
            xaxis_title="", margin=dict(t=30, b=10), height=440,
        )
        return fig

    # Recorte para uma única categoria de RUSH ("Sim" ou "Não")
    subset = df[df["RUSH"] == rush_opcao]
    df_media = pd.DataFrame(_media_por_etapa(subset))
    fig = px.bar(
        df_media, x="Etapa", y="Média", color="Etapa",
        color_discrete_map=CORES_ETAPAS, text="Rotulo", custom_data=["Qtd"],
    )
    fig.update_traces(
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Média: %{text} dias<br>Remessas consideradas: %{customdata[0]}<extra></extra>",
    )
    fig.update_layout(showlegend=False, yaxis_title="Média de dias", yaxis_tickformat="d",
                       xaxis_title="", margin=dict(t=30, b=10), height=420,
                       title=f"Recorte: RUSH = {rush_opcao}")
    return fig


# ==============================================================================
# 5. GRÁFICO 2 — BARRAS EMPILHADAS POR REMESSA (com destaque de RUSH)
# ==============================================================================
def criar_grafico_remessas(df_pagina: pd.DataFrame) -> go.Figure:
    """
    Colunas verticais empilhadas: uma coluna por remessa (já paginado
    externamente — ver `main()`). Remessas marcadas como RUSH ("Sim") ganham
    o número em verde e negrito no eixo x.
    """
    df_g = df_pagina

    tem_rush = "RUSH" in df_g.columns
    if tem_rush:
        eh_rush = (df_g["RUSH"] == "Sim").tolist()
    else:
        eh_rush = [False] * len(df_g)

    # Remessas RUSH têm o próprio número (não a data) pintado de verde e em
    # negrito no eixo x. Plotly interpreta um pequeno subconjunto de HTML em
    # rótulos de eixo (<b>, <span style="color:...">, <br>).
    rotulos = []
    for rem, dt, rush in zip(df_g["REMESSA"], df_g["DATA_REPORTE"], eh_rush):
        numero = f'<span style="color:green"><b>{rem}</b></span>' if rush else f"{rem}"
        rotulos.append(f"{numero}<br>{dt:%d/%m/%Y}")

    fig = go.Figure()
    for col, nome in ETAPAS.items():
        dias = df_g[col].fillna(0).clip(lower=0).astype(int)
        fig.add_trace(
            go.Bar(
                x=rotulos,
                y=dias,
                name=nome,
                marker_color=CORES_ETAPAS[nome],
                text=[str(d) if d > 0 else "" for d in dias],
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(color="white", size=13),
                hovertemplate="%{x}<br>" + nome + "<br>Dias: %{y:.0f}<extra></extra>",
            )
        )

    # Total no topo de cada coluna
    totais = df_g["Dias_total"].astype(int)
    fig.add_trace(
        go.Scatter(
            x=rotulos,
            y=totais,
            mode="text",
            text=[f"<b>{t} dias</b>" for t in totais],
            textposition="top center",
            showlegend=False,
            hoverinfo="skip",
        )
    )

    fig.update_layout(
        barmode="stack",
        height=500,
        legend_title="Etapa",
        legend=dict(
            orientation="v",
            traceorder="normal",   # Envio → Notificação no topo, na ordem das etapas
            yanchor="top", y=1,
            xanchor="left", x=1.02,
        ),
        margin=dict(t=40, b=10, l=10, r=10),
        yaxis=dict(visible=False, range=[0, (totais.max() if len(totais) else 1) * 1.15]),
        xaxis=dict(type="category", categoryorder="array", categoryarray=rotulos, title=""),
    )
    return fig


# ==============================================================================
# 6. GRÁFICO 3 — LINHA: TOTAL DE DIAS POR REMESSA AO LONGO DO TEMPO
# ==============================================================================
def criar_grafico_linha_tempo(
    df: pd.DataFrame, limite_atencao: int, limite_critico: int
) -> go.Figure:
    """
    Gráfico de linha com o total de dias (Dias_total) de cada remessa
    finalizada, posicionada no eixo x pela data de finalização (DATA_REPORTE).

    - Cada ponto é uma remessa (hover mostra número, data e total de dias).
    - Remessas RUSH aparecem como marcador verde (losango).
    - Linha tracejada com a média móvel das últimas 5 remessas (tendência).
    - Linhas horizontais com os limites de atenção e crítico da barra lateral.
    """
    df_l = df.dropna(subset=["DATA_REPORTE"]).sort_values("DATA_REPORTE").copy()
    df_l["Dias_total"] = df_l["Dias_total"].astype(int)
    df_l["Media_movel"] = df_l["Dias_total"].rolling(window=5, min_periods=1).mean()

    eh_rush = (df_l["RUSH"] == "Sim") if "RUSH" in df_l.columns else pd.Series(False, index=df_l.index)
    rush_txt = np.where(eh_rush, "Sim", "Não")
    customdata = np.column_stack([df_l["REMESSA"].astype(str), rush_txt])

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df_l["DATA_REPORTE"],
            y=df_l["Dias_total"],
            mode="lines+markers",
            name="Total de dias",
            line=dict(color="#4C78A8", width=2),
            marker=dict(
                size=np.where(eh_rush, 11, 7),
                color=np.where(eh_rush, "green", "#4C78A8"),
                symbol=np.where(eh_rush, "diamond", "circle"),
                line=dict(color="white", width=1),
            ),
            customdata=customdata,
            hovertemplate=(
                "<b>Remessa %{customdata[0]}</b><br>"
                "Finalização: %{x|%d/%m/%Y}<br>"
                "Total: %{y} dias<br>"
                "RUSH: %{customdata[1]}<extra></extra>"
            ),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=df_l["DATA_REPORTE"],
            y=df_l["Media_movel"],
            mode="lines",
            name="Média móvel (5 remessas)",
            line=dict(color="#F58518", width=2, dash="dash"),
            hovertemplate="Média móvel: %{y:.0f} dias<extra></extra>",
        )
    )

    fig.add_hline(
        y=limite_atencao, line_dash="dot", line_color="#E8B300",
        annotation_text=f"Atenção ({limite_atencao} dias)", annotation_position="top left",
    )
    fig.add_hline(
        y=limite_critico, line_dash="dot", line_color="#E45756",
        annotation_text=f"Crítico ({limite_critico} dias)", annotation_position="top left",
    )

    fig.update_layout(
        height=450,
        xaxis_title="Data de finalização (DATA_REPORTE)",
        yaxis_title="Total de dias",
        yaxis_tickformat="d",
        xaxis_tickformat="%d/%m/%Y",
        hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, b=10),
    )
    return fig


# ==============================================================================
# 7. REMESSAS EM ANDAMENTO (NÃO FINALIZADAS) — ACOMPANHAMENTO POR ETAPA
# ==============================================================================
QTD_POR_PAGINA_ANDAMENTO = 10


def preparar_remessas_em_andamento(df: pd.DataFrame) -> pd.DataFrame:
    """
    Seleciona as remessas NÃO finalizadas (DATA_REPORTE vazia) e monta os dias
    por etapa usando SOMENTE etapas com as duas datas preenchidas (diferença já
    calculada). Etapa com data ausente não é exibida (fica em 0). Remessas sem
    nenhuma etapa calculável são descartadas.
    """
    base = df[df["DATA_REPORTE"].isna()].copy()

    registros = pd.DataFrame(
        {
            "REMESSA": base["REMESSA"],
            "RUSH": base["RUSH"] if "RUSH" in base.columns else np.nan,
            "DATA_ENVIO": base["DATA_ENVIO"],
        }
    )
    for col, nome in ETAPAS.items():
        registros[nome] = base[col].fillna(0).clip(lower=0).astype(int)

    registros["Dias_total"] = registros[list(ETAPAS.values())].sum(axis=1).astype(int)
    registros = registros[registros["Dias_total"] > 0]
    return registros.sort_values("Dias_total", ascending=False).reset_index(drop=True)


def criar_grafico_andamento(df_pagina: pd.DataFrame) -> go.Figure:
    """
    Colunas empilhadas por remessa em andamento, com os dias de cada etapa já
    calculada (datas de início e fim preenchidas). Número da remessa RUSH em verde.
    """
    rotulos = []
    for rem, dt, rush in zip(df_pagina["REMESSA"], df_pagina["DATA_ENVIO"], df_pagina["RUSH"]):
        numero = f'<span style="color:green"><b>{rem}</b></span>' if rush == "Sim" else f"{rem}"
        data_txt = f"{dt:%d/%m/%Y}" if pd.notna(dt) else "s/ data"
        rotulos.append(f"{numero}<br>{data_txt}")

    fig = go.Figure()
    for nome in ETAPAS.values():
        dias = df_pagina[nome].astype(int)
        fig.add_trace(
            go.Bar(
                x=rotulos,
                y=dias,
                name=nome,
                marker_color=CORES_ETAPAS[nome],
                text=[str(d) if d > 0 else "" for d in dias],
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(color="white", size=13),
                hovertemplate="%{x}<br>" + nome + "<br>Dias: %{y:.0f}<extra></extra>",
            )
        )

    totais = df_pagina["Dias_total"].astype(int)
    fig.add_trace(
        go.Scatter(
            x=rotulos,
            y=totais,
            mode="text",
            text=[f"<b>{t} dias</b>" for t in totais],
            textposition="top center",
            showlegend=False,
            hoverinfo="skip",
        )
    )

    fig.update_layout(
        barmode="stack",
        height=500,
        legend_title="Etapa",
        legend=dict(
            orientation="v",
            traceorder="normal",   # Envio → Notificação no topo, na ordem das etapas
            yanchor="top", y=1,
            xanchor="left", x=1.02,
        ),
        margin=dict(t=40, b=10, l=10, r=10),
        yaxis=dict(visible=False, range=[0, (totais.max() if len(totais) else 1) * 1.15 or 1]),
        xaxis=dict(type="category", categoryorder="array", categoryarray=rotulos, title=""),
    )
    return fig


def exibir_remessas_em_andamento(df: pd.DataFrame) -> None:
    """Renderiza a seção 'Remessas em Andamento' (KPIs, paginação e gráfico)."""
    st.markdown("---")
    st.subheader("🚧 Remessas em Andamento")
    st.caption(
        "Remessas ainda **não finalizadas** (DATA_REPORTE vazia). Cada etapa só aparece quando as "
        "duas datas dela estão preenchidas; etapas sem data não são exibidas nem estimadas. "
        "Número da remessa em **verde** = RUSH. Ordenadas das que têm mais dias para as com menos."
    )

    qtd_nao_finalizadas = int(df["DATA_REPORTE"].isna().sum())
    df_and = preparar_remessas_em_andamento(df)

    if df_and.empty:
        st.info("Nenhuma remessa em andamento com etapas calculáveis nos filtros atuais.")
        return

    c1, c2 = st.columns(2)
    c1.metric("🚧 Em andamento (exibidas)", f"{len(df_and)}")
    c2.metric("📊 Média de dias (etapas preenchidas)", f"{df_and['Dias_total'].mean():.0f}")

    ocultas = qtd_nao_finalizadas - len(df_and)
    if ocultas > 0:
        st.caption(f"ℹ️ {ocultas} remessa(s) não finalizada(s) ficaram de fora por não ter nenhuma etapa com as duas datas preenchidas.")

    total = len(df_and)
    st.session_state.setdefault("pagina_andamento", 0)
    max_pagina = max((total - 1) // QTD_POR_PAGINA_ANDAMENTO, 0)
    st.session_state["pagina_andamento"] = min(st.session_state["pagina_andamento"], max_pagina)
    pagina = st.session_state["pagina_andamento"]
    inicio = pagina * QTD_POR_PAGINA_ANDAMENTO
    fim = inicio + QTD_POR_PAGINA_ANDAMENTO

    col_prev, col_info, col_next = st.columns([1, 3, 1])
    with col_prev:
        if st.button("◀ Anteriores", disabled=(pagina == 0), key="btn_andamento_anterior"):
            st.session_state["pagina_andamento"] = max(pagina - 1, 0)
            st.rerun()
    with col_info:
        st.markdown(
            f"<div style='text-align:center;'>Exibindo remessas "
            f"<b>{inicio + 1}–{min(fim, total)}</b> de <b>{total}</b></div>",
            unsafe_allow_html=True,
        )
    with col_next:
        if st.button("Próximas ▶", disabled=(fim >= total), key="btn_andamento_seguinte"):
            st.session_state["pagina_andamento"] = min(pagina + 1, max_pagina)
            st.rerun()

    st.plotly_chart(criar_grafico_andamento(df_and.iloc[inicio:fim]), width="stretch")


# ==============================================================================
# 8. INCONSISTÊNCIAS
# ==============================================================================
def identificar_inconsistencias(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna apenas as remessas com pelo menos uma diferença negativa."""
    df_inc = df[df["Possui_inconsistencia"]].copy()
    colunas = ["REMESSA", "PROJETO", "RESPONSÁVEL"] + COLS_DIFF + ["Dias_total"]
    colunas = [c for c in colunas if c in df_inc.columns]
    df_inc = df_inc[colunas].rename(columns={**ETAPAS, "Dias_total": "Dias totais"})
    return df_inc


# ==============================================================================
# 9. EXPORTAÇÃO
# ==============================================================================
def exportar_dados(df: pd.DataFrame) -> None:
    """Cria botões de download (Excel e CSV) com os dados filtrados."""
    col1, col2 = st.columns(2)

    csv_bytes = df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
    col1.download_button(
        "⬇️ Baixar CSV (dados filtrados)",
        data=csv_bytes,
        file_name=f"remessas_filtradas_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        width='stretch',
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Remessas")
    col2.download_button(
        "⬇️ Baixar Excel (dados filtrados)",
        data=buffer.getvalue(),
        file_name=f"remessas_filtradas_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width='stretch',
    )


# ==============================================================================
# APLICAÇÃO PRINCIPAL
# ==============================================================================
def main():
    st.title("📦 Dashboard de Tempo de Processamento de Remessas")
    st.caption(
        "Análise interativa do tempo gasto em cada etapa do processo: "
        "Envio → Notificação → Pagamento → Início da Preparação → Finalização"
    )

    # --- Carga e preparação ---
    df_raw = carregar_df()
    df_prep = preparar_dados(df_raw)
    df_calc = calcular_diferencas(df_prep)

    # --- Filtro fixo: apenas remessas do TIPO "RC" ---
    if "TIPO" in df_calc.columns:
        df_calc = df_calc[df_calc["TIPO"] == "RC"].copy()
    if df_calc.empty:
        st.warning('Nenhuma remessa com TIPO == "RC" encontrada na base carregada.')
        return
    st.caption('🔎 Exibindo apenas remessas com **TIPO = "RC"**.')

    # --- Filtros (inclui limites de classificação) ---
    df_filtrado, limite_atencao, limite_critico = aplicar_filtros(df_calc)
    df_filtrado = classificar_remessas(df_filtrado, limite_atencao, limite_critico)

    if df_filtrado.empty:
        st.warning("Nenhuma remessa encontrada para os filtros selecionados.")
        return

    # --- KPIs ---
    st.markdown("---")
    st.subheader("📊 Indicadores")
    criar_kpis(df_filtrado)

    # --- Gráfico 1: Média por etapa (com recorte por RUSH) ---
    st.markdown("---")
    st.subheader("⏱️ Média de Dias por Etapa")
    st.caption(
        "Considera apenas remessas com DATA_REPORTE preenchida (finalizadas). "
        f"A(s) remessa(s) {', '.join(sorted(REMESSAS_EXCLUIDAS_MEDIA))} "
        "é(são) excluída(s) deste cálculo."
    )

    if "RUSH" in df_filtrado.columns:
        rush_opcao = st.radio(
            "Categoria RUSH",
            options=["Todas (comparar Sim x Não)", "Sim", "Não"],
            horizontal=True,
            key="rush_opcao_media",
        )
    else:
        rush_opcao = "Todas (comparar Sim x Não)"
        st.info('Coluna "RUSH" não encontrada na base — exibindo média geral, sem recorte.')

    qtd_finalizadas = (
        df_filtrado["DATA_REPORTE"].notna().sum() if "DATA_REPORTE" in df_filtrado.columns else 0
    )
    if qtd_finalizadas == 0:
        st.warning(
            "Nenhuma remessa com DATA_REPORTE preenchida nos filtros atuais — "
            "a média por etapa não pode ser calculada."
        )
    else:
        st.plotly_chart(criar_grafico_media_etapas(df_filtrado, rush_opcao), width='stretch')

    # --- Gráfico 2: Barras empilhadas por remessa (com paginação e clique) ---
    st.markdown("---")
    st.subheader("📈 Tempo de Processamento por Remessa")
    st.caption(
        "Número da remessa em **verde** no eixo x = remessa marcada como RUSH."
    )

    df_ordenado = df_filtrado.dropna(subset=["DATA_REPORTE"]).sort_values(
        "DATA_REPORTE", ascending=False
    )
    total_remessas = len(df_ordenado)

    if total_remessas == 0:
        st.info("Nenhuma remessa finalizada (com DATA_REPORTE) para exibir.")
    else:
        st.session_state.setdefault("pagina_remessas", 0)
        max_pagina = max((total_remessas - 1) // QTD_POR_PAGINA, 0)
        # Garante que a página armazenada continua válida após mudança nos filtros
        st.session_state["pagina_remessas"] = min(st.session_state["pagina_remessas"], max_pagina)
        pagina = st.session_state["pagina_remessas"]

        inicio = pagina * QTD_POR_PAGINA
        fim = inicio + QTD_POR_PAGINA
        df_pagina = df_ordenado.iloc[inicio:fim]

        col_prev, col_info, col_next = st.columns([1, 3, 1])
        with col_prev:
            if st.button("◀ Mais recentes", disabled=(pagina == 0), key="btn_pagina_anterior"):
                st.session_state["pagina_remessas"] = max(pagina - 1, 0)
                st.rerun()
        with col_info:
            st.markdown(
                f"<div style='text-align:center;'>Exibindo remessas "
                f"<b>{inicio + 1}–{min(fim, total_remessas)}</b> de <b>{total_remessas}</b></div>",
                unsafe_allow_html=True,
            )
        with col_next:
            if st.button("Anteriores ▶", disabled=(fim >= total_remessas), key="btn_pagina_seguinte"):
                st.session_state["pagina_remessas"] = min(pagina + 1, max_pagina)
                st.rerun()

        fig_remessas = criar_grafico_remessas(df_pagina)

        st.plotly_chart(fig_remessas, width="stretch")

    # --- Gráfico 3: Linha com total de dias por remessa ao longo do tempo ---
    st.markdown("---")
    st.subheader("📉 Total de Dias por Remessa ao Longo do Tempo")
    st.caption(
        "Cada ponto é uma remessa finalizada, posicionada pela data de finalização "
        "(DATA_REPORTE). Losango **verde** = RUSH. A linha tracejada laranja mostra a "
        "média móvel das últimas 5 remessas; as linhas pontilhadas são os limites de "
        "atenção e crítico definidos na barra lateral."
    )
    if total_remessas == 0:
        st.info("Nenhuma remessa finalizada (com DATA_REPORTE) para exibir.")
    else:
        st.plotly_chart(
            criar_grafico_linha_tempo(df_filtrado, limite_atencao, limite_critico),
            width="stretch",
        )

    # --- Remessas em andamento (não finalizadas) ---
    exibir_remessas_em_andamento(df_filtrado)

    # --- Planilha completa ---
    st.markdown("---")
    st.subheader("🗂️ Planilha Completa de Dados")
    st.caption(f"{len(df_filtrado)} remessa(s) conforme os filtros selecionados.")
    cols_aux = [c for c in df_filtrado.columns if c.endswith("__inconsistente")] + [
        f"etapa{i}_incompleta" for i in range(1, 5)
    ]
    df_completo = df_filtrado.drop(columns=cols_aux, errors="ignore").rename(
        columns={**ETAPAS, "Dias_total": "Dias totais"}
    )
    df_completo = formatar_datas_para_exibicao(df_completo)
    st.dataframe(df_completo, width='stretch', hide_index=True, height=500)

    # --- Inconsistências ---
    st.markdown("---")
    st.subheader("⚠️ Inconsistências de Datas (diferenças negativas)")
    df_inconsist = identificar_inconsistencias(df_filtrado)
    if df_inconsist.empty:
        st.success("Nenhuma inconsistência de datas encontrada nos dados filtrados.")
    else:
        st.error(f"{len(df_inconsist)} remessa(s) com diferenças negativas — possível inconsistência de datas.")
        st.dataframe(df_inconsist, width='stretch', hide_index=True)

    # --- Processo incompleto ---
    df_incompletas = df_filtrado[df_filtrado["Processo_incompleto"]]
    if not df_incompletas.empty:
        st.warning(
            f"ℹ️ {len(df_incompletas)} remessa(s) possuem etapas incompletas "
            f"(datas ausentes) e não entram no cálculo da média daquela etapa."
        )

    # --- Exportação ---
    st.markdown("---")
    st.subheader("💾 Exportar Dados Filtrados")
    exportar_dados(df_filtrado)


if __name__ == "__main__":
    main()
