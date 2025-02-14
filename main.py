import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Nome do arquivo para armazenar os registros
FILE_NAME = "registro_ponto.xlsx"

def carregar_dados():
    """Carrega os dados do arquivo, ou cria um novo se não existir."""
    if os.path.exists(FILE_NAME):
        df = pd.read_excel(FILE_NAME)
        # Garantir que as colunas de hora sejam do tipo string ou datetime
        df['Hora Entrada'] = pd.to_datetime(df['Hora Entrada'], errors='coerce').dt.strftime('%H:%M:%S')
        df['Hora Saída'] = pd.to_datetime(df['Hora Saída'], errors='coerce').dt.strftime('%H:%M:%S')
        return df
    else:
        return pd.DataFrame(columns=["Data", "Hora Entrada", "Hora Saída", "Horas Trabalhadas", "Valor Recebido", "Mês"])

def salvar_dados(df):
    """Salva os dados no arquivo."""
    # Garantir que as colunas de hora estejam no formato correto antes de salvar
    df['Hora Entrada'] = pd.to_datetime(df['Hora Entrada'], errors='coerce').dt.strftime('%H:%M:%S')
    df['Hora Saída'] = pd.to_datetime(df['Hora Saída'], errors='coerce').dt.strftime('%H:%M:%S')
    df.to_excel(FILE_NAME, index=False)


def bater_ponto():
    """Registra um novo ponto no arquivo."""
    df = carregar_dados()
    agora = datetime.now()
    data = agora.strftime("%Y-%m-%d")
    hora = agora.strftime("%H:%M:%S")
    mes = agora.strftime("%Y-%m")
    
    if not df.empty and df.iloc[-1]['Data'] == data and pd.isna(df.iloc[-1]['Hora Saída']):
        # Se já tem um registro de entrada no mesmo dia, registra a saída
        df.at[df.index[-1], 'Hora Saída'] = hora
        entrada = datetime.strptime(df.iloc[-1]['Hora Entrada'], "%H:%M:%S")
        saida = datetime.strptime(hora, "%H:%M:%S")
        horas_trabalhadas = (saida - entrada).total_seconds() / 3600 - 1  # Desconta 1h de almoço
        horas_trabalhadas = max(horas_trabalhadas, 0)  # Garante que não seja negativo
        df.at[df.index[-1], 'Horas Trabalhadas'] = round(horas_trabalhadas, 2)
        df.at[df.index[-1], 'Valor Recebido'] = round(horas_trabalhadas * 30, 2)
    else:
        # Novo registro de entrada
        novo_registro = pd.DataFrame([{ "Data": data, "Hora Entrada": hora, "Hora Saída": None, "Horas Trabalhadas": None, "Valor Recebido": None, "Mês": mes }])
        df = pd.concat([df, novo_registro], ignore_index=True)
    
    salvar_dados(df)
    return df

def calcular_totais(mes):
    """Calcula total de horas e valor recebido no mês selecionado."""
    df = carregar_dados()
    df_mes = df[df["Mês"] == mes]
    total_horas = df_mes["Horas Trabalhadas"].sum()
    total_valor = df_mes["Valor Recebido"].sum()
    return round(total_horas, 2), round(total_valor, 2)

# Interface Streamlit
st.title("📌 Sistema de Ponto")

if st.button("🕒 Bater Ponto"):
    df_atualizado = bater_ponto()
    st.success("Ponto registrado com sucesso!")

# Seleção de mês
st.subheader("📅 Selecione o mês para visualizar")
df = carregar_dados()
meses = df["Mês"].dropna().unique()
mes_selecionado = st.selectbox("Escolha o mês", meses if len(meses) > 0 else [datetime.now().strftime("%Y-%m")])

total_horas, total_valor = calcular_totais(mes_selecionado)
st.metric("Total de Horas Trabalhadas", f"{total_horas} h")
st.metric("Total a Receber", f"R$ {total_valor}")

st.subheader("📄 Registros de Ponto")
st.dataframe(df[df["Mês"] == mes_selecionado])

# Download da planilha
if not df.empty:
    df_filtrado = df[df["Mês"] == mes_selecionado]
    df_filtrado.to_excel("horas_trabalhadas.xlsx", index=False)
    with open("horas_trabalhadas.xlsx", "rb") as file:
        st.download_button("📥 Baixar Planilha do Mês Selecionado", data=file, file_name=f"horas_trabalhadas_{mes_selecionado}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
