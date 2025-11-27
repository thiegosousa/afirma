import re
import io
import pandas as pd
import streamlit as st

DEFAULT_FILE_PATH = "modelo_planilha_nova.xlsx"  # Atualizado para seu arquivo real
st.set_page_config(page_title="Transpor Planilha AFIRMASUS - CORRETO", layout="centered")

st.title("Planilha AFIRMASUS → Formato Final 100% Correto")
st.caption("Agência, Dígito, Conta e Dígito1 agora reconhecidos perfeitamente")

# === 14 COLUNAS FINAIS ===
COLUNAS_FINAIS = [
    "Nome completo",
    "CPF",
    "Data de Nascimento",
    "Tipo",
    "Graduação",
    "Número",
    "Instituição Bancária",
    "Agência Bancária (sem dígito)",
    "Digito",
    "Número da Conta Corrente Nominal (sem dígito)",
    "Dígito1",
    "Instituição de ensino superior",
    "campus",
    "Nome da(o) Tutura(or)"
]

# === COLUNAS A IGNORAR (Nomes completos das mães) ===
COLUNAS_IGNORAR = [
    f"Estudante {i}) Nome completo da mãe" for i in range(1, 11)
] + ["Orientadora(or) de Serviço) Nome completo da mãe"]

# ============================================================
# FUNÇÃO PRINCIPAL DE EXTRAÇÃO DE LABEL (CORRIGIDA E ROBUSTA)
# ============================================================
def extrair_label(coluna: str) -> str:
    texto = str(coluna).replace("\xa0", " ").strip()
    txt = texto.lower()

    # Ignora nome da mãe
    if "nome completo da mãe" in txt:
        return ""

    # ==================== DADOS BANCÁRIOS - ESTUDANTES ====================
    if re.search(r"Estudante\s*\d*\s*\)\s*Agência Bancária\s*\(sem dígito\)", texto):
        return "Agência Bancária (sem dígito)"
    if re.search(r"Estudante\s*\d*\s*\)\s*Dígito$", texto):
        return "Digito"
    if re.search(r"Estudante\s*\d*\s*\)\s*Número da Conta Corrente Nominal\s*\(sem dígito\)", texto):
        return "Número da Conta Corrente Nominal (sem dígito)"
    if re.search(r"Estudante\s*\d*\s*\)\s*Dígito1$", texto):
        return "Dígito1"

    # ==================== DADOS BANCÁRIOS - ORIENTADOR ====================
    if "Orientadora(or) de Serviço) Agência Bancária (sem dígito)" in texto:
        return "Agência Bancária (sem dígito)"
    if "Orientadora(or) de Serviço) Dígito" in texto:
        return "Digito"
    if "Orientadora(or) de Serviço) Número da Conta Corrente Norminal (sem dígito)" in texto:
        return "Número da Conta Corrente Nominal (sem dígito)"
    if "Orientadora(or) de Serviço) Dígito1" in texto:
        return "Dígito1"

    # ==================== OUTROS CAMPOS PADRÃO ====================
    if "cpf" in txt:
        return "CPF"
    if "nome completo" in txt and "mãe" not in txt:
        return "Nome completo"
    if "nascimento" in txt:
        return "Data de Nascimento"
    if "instituição bancária" in txt:
        return "Instituição Bancária"
    if "graduação" in txt or "nível de formação" in txt:
        return "Graduação"
    if "instituição de ensino superior" in txt:
        return "Instituição de ensino superior"
    if txt == "campus":
        return "campus"
    if "tutora" in txt or "tutor" in txt:
        return "Nome da(o) Tutura(or)"

    return ""

# ============================================================
# DETECTAR GRUPOS (Estudantes 1 a 10 + Orientador)
# ============================================================
def detect_groups(columns):
    grupos = {}
    fixas = []

    for col in columns:
        if col in COLUNAS_IGNORAR:
            continue

        label = extrair_label(col)
        if not label:
            continue

        col_str = str(col)

        # Estudantes
        match_est = re.search(r"Estudante\s+(\d+)\)", col_str)
        if match_est:
            num = match_est.group(1)
            key = f"Estudante {num}"
            grupos.setdefault(key, {})[label] = col

        # Orientador
        elif "Orientadora(or)" in col_str:
            grupos.setdefault("Orientador", {})[label] = col

        else:
            fixas.append(col)

    return grupos, fixas

# ============================================================
# FUNÇÕES DE FORMATAÇÃO
# ============================================================
def limpar_num(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\D", "", str(v))

def formatar_cpf(v):
    cpf = limpar_num(v).zfill(11)[-11:]
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if len(cpf) == 11 else ""

def formatar_data(v):
    if pd.isna(v):
        return ""
    try:
        return pd.to_datetime(v, dayfirst=True, errors='coerce').strftime("%d/%m/%Y")
    except:
        return str(v) if pd.notna(v) else ""

# ============================================================
# TRANSFORMAÇÃO FINAL
# ============================================================
def transformar(df, grupos, fixas):
    linhas = []
    estudantes = sorted([k for k in grupos if k.startswith("Estudante")], key=lambda x: int(x.split()[-1]))

    for _, row in df.iterrows():
        base = {}
        for col in fixas:
            label = extrair_label(col)
            if label in COLUNAS_FINAIS:
                base[label] = row.get(col, "")

        # === Processa cada Estudante ===
        for key in estudantes:
            campos = grupos[key]
            if "Nome completo" not in campos or pd.isna(row[campos["Nome completo"]]) or str(row[campos["Nome completo"]]).strip() == "":
                continue

            reg = base.copy()
            reg["Tipo"] = "Estudante"
            reg["Número"] = key

            for label, col_orig in campos.items():
                val = row.get(col_orig, "")

                if label == "CPF":
                    val = formatar_cpf(val)
                elif label == "Data de Nascimento":
                    val = formatar_data(val)
                elif label in ["Agência Bancária (sem dígito)", "Digito", "Número da Conta Corrente Nominal (sem dígito)", "Dígito1"]:
                    val = limpar_num(val)

                reg[label] = val

            linhas.append(reg)

        # === Processa Orientador ===
        if "Orientador" in grupos:
            campos = grupos["Orientador"]
            if "Nome completo" in campos and pd.notna(row[campos["Nome completo"]]) and str(row[campos["Nome completo"]]).strip():
                reg = base.copy()
                reg["Tipo"] = "Orientador"
                reg["Número"] = "Orientador"

                for label, col_orig in campos.items():
                    val = row.get(col_orig, "")

                    if label == "CPF":
                        val = formatar_cpf(val)
                    elif label == "Data de Nascimento":
                        val = formatar_data(val)
                    elif label in ["Agência Bancária (sem dígito)", "Digito", "Número da Conta Corrente Nominal (sem dígito)", "Dígito1"]:
                        val = limpar_num(val)

                    reg[label] = val

                linhas.append(reg)

    df_final = pd.DataFrame(linhas)
    for col in COLUNAS_FINAIS:
        df_final[col] = df_final.get(col, "")

    return df_final[COLUNAS_FINAIS]

# ============================================================
# INTERFACE STREAMLIT
# ============================================================
st.sidebar.header("Upload da Planilha")
use_default = st.sidebar.checkbox("Usar modelo_planilha_nova.xlsx (padrão)", value=True)
uploaded_file = st.sidebar.file_uploader("Ou faça upload do seu arquivo", type=["xlsx"])

if use_default and not uploaded_file:
    uploaded_file = DEFAULT_FILE_PATH
    st.info(f"Usando arquivo padrão: `{DEFAULT_FILE_PATH}`")

if not uploaded_file:
    st.info("Aguardando upload do arquivo...")
    st.stop()

try:
    df = pd.read_excel(uploaded_file, dtype=str)
    df = df.drop(columns=[c for c in COLUNAS_IGNORAR if c in df.columns], errors="ignore")
    st.success(f"Arquivo carregado com sucesso: {df.shape[0]} linhas × {df.shape[1]} colunas")
except Exception as e:
    st.error(f"Erro ao carregar: {e}")
    st.stop()

grupos, fixas = detect_groups(df.columns)

if st.button("GERAR PLANILHA FINAL CORRIGIDA", type="primary"):
    with st.spinner("Transformando planilha..."):
        resultado = transformar(df, grupos, fixas)

    st.success(f"PRONTO! {len(resultado)} registros gerados (Estudantes + Orientadores)")
    st.dataframe(resultado, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        resultado.to_excel(writer, index=False, sheet_name="Pessoas")
    buffer.seek(0)

    st.download_button(
        label="BAIXAR PLANILHA FINAL (CORRETA)",
        data=buffer,
        file_name="AFIRMASUS_FINAL_CORRIGIDO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
