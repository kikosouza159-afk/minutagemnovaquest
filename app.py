from flask import Flask, jsonify, render_template
import pandas as pd
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(BASE_DIR, "Minutagem_Locator_NovaQuest_Aut.xlsx")


def arquivo_existe(caminho):
    return os.path.exists(caminho)


def parse_texto(valor, padrao=""):
    if pd.isna(valor):
        return padrao
    return str(valor).strip()


def parse_inteiro(valor, padrao=0):
    if pd.isna(valor) or valor == "":
        return padrao
    try:
        if isinstance(valor, str):
            valor = valor.strip().replace(".", "").replace(",", ".")
        return int(float(valor))
    except Exception:
        return padrao


def parse_decimal(valor, padrao=0.0):
    if pd.isna(valor) or valor == "":
        return padrao
    try:
        if isinstance(valor, str):
            valor = valor.strip().replace(".", "").replace(",", ".")
        return float(valor)
    except Exception:
        return padrao


def normalizar_data(valor):
    if pd.isna(valor) or valor == "":
        return ""
    try:
        return pd.to_datetime(valor).strftime("%Y-%m-%d")
    except Exception:
        return str(valor).strip()


@app.route("/")
def home():
    return render_template("dashboard_bilhetagem.html")


@app.route("/dados")
def dados():
    if not arquivo_existe(ARQUIVO_EXCEL):
        return jsonify({
            "erro": True,
            "mensagem": f"Arquivo não encontrado: {ARQUIVO_EXCEL}"
        }), 404

    try:
        df = pd.read_excel(ARQUIVO_EXCEL)
        df.columns = [str(col).strip() for col in df.columns]
        df = df.fillna("")

        mapa_colunas = {
            "Date": "Date",
            "date": "Date",
            "Hour": "Hour",
            "hour": "Hour",
            "CampaignId": "CampaignId",
            "CampaignID": "CampaignId",
            "Flow": "Flow",
            "flow": "Flow",
            "TotalCalls": "TotalCalls",
            "TotalTalkingTime": "TotalTalkingTime",
            "TotalBillingTime": "TotalBillingTime",
            "TotalValue": "TotalValue",
        }

        df = df.rename(columns={c: mapa_colunas.get(c, c) for c in df.columns})

        registros = []
        for _, row in df.iterrows():
            registros.append({
                "date": normalizar_data(row.get("Date", "")),
                "hour": parse_inteiro(row.get("Hour", 0)),
                "campaignId": parse_inteiro(row.get("CampaignId", 0)),
                "flow": parse_texto(row.get("Flow", "")),
                "totalCalls": parse_inteiro(row.get("TotalCalls", 0)),
                "totalTalkingTime": parse_decimal(row.get("TotalTalkingTime", 0)),
                "totalBillingTime": parse_decimal(row.get("TotalBillingTime", 0)),
                "totalValue": parse_decimal(row.get("TotalValue", 0)),
            })

        return jsonify(registros)

    except Exception as e:
        return jsonify({
            "erro": True,
            "mensagem": f"Erro ao ler o Excel: {str(e)}"
        }), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)