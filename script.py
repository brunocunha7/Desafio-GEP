import requests
from pathlib import Path
import json
import pandas as pd
from datetime import date, timedelta
import re
import logging

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def get_token() -> str:
    api_route = "https://63p7r2qck2.execute-api.us-east-1.amazonaws.com/Prod/token"
    group_id = "11411183-c06e-4690-9537-67a40c1df2ca"
    report_id = "36f8f9aa-cf5a-4bd2-b09f-87b1b06ac1eb"

    url = f"{api_route}/{group_id}/{report_id}"
    logging.info("Requisitando token em %s", url)
    r = requests.get(url, timeout=30)
    try:
        r.raise_for_status()
    except Exception as e:
        logging.error("Failed to get token: %s", e)
        raise
    data = r.json()

    embed_url = data.get("EmbedUrl")
    embed_token = data.get("Token")

    logging.debug("EmbedUrl: %s", embed_url)
    token = "EmbedToken " + embed_token
    return token

def atualizar_payload(caminho_json: str, novo_produto: str) -> None:
    logging.info("Atualizando payload para produto: %s", novo_produto)
    with open(caminho_json, "r", encoding="utf-8") as f:
        payload = json.load(f)

    comandos = payload["executeSemanticQueryRequest"]["queries"][0]["Query"]["Commands"]

    #Atualiza o valor do filtro PRODUCTO
    in_cond = comandos[0]["SemanticQueryDataShapeCommand"]["Query"]["Where"][1]["Condition"]["In"]
    in_cond["Values"][0][0]["Literal"]["Value"] = f"'{novo_produto}'"

    #Atualiza a data final para o dia de ontem (verificar a atualização no site)
    ontem = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    comandos[0]["SemanticQueryDataShapeCommand"]["Query"]["Where"][0]["Condition"]["And"]["Right"]["Comparison"]["Right"]["Literal"]["Value"] = (
        f"datetime'{ontem}T00:00:00'"
    )

    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    logging.debug("Payload salvo em %s", caminho_json)

def download_xlsx(out_path: str = "export.xlsx") -> Path:
    logging.info("Iniciando download para %s", out_path)
    # cria pasta produtos, caso não exista
    Path("produtos").mkdir(parents=True, exist_ok=True)

    URL = "https://wabi-south-central-us-redirect.analysis.windows.net/export/xlsx"
    token = get_token()

    with open("payload.json", "r", encoding="utf-8") as f:
        payload = json.load(f)

    headers = {
        "Authorization": token,
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://app.powerbi.com",
        "Referer": "https://app.powerbi.com/",
        "X-PowerBI-HostEnv": "Embed for Customers",
        "User-Agent": "Mozilla/5.0",
    }

    out = Path(out_path)

    with requests.Session() as s:
        r = s.post(URL, headers=headers, json=payload, timeout=180)
        try:
            r.raise_for_status()
        except Exception as e:
            logging.error("Download falhou: %s", e)
            logging.debug("Resposta: %s", getattr(r, "text", "")[:2000])
            raise

        ctype = (r.headers.get("content-type") or "").lower()

        if "spreadsheetml.sheet" not in ctype:
            logging.error("Content-Type inesperado: %s", ctype)
            logging.debug("Resposta (primeiros 2000 bytes): %s", r.content[:2000])
            raise RuntimeError("Não retornou XLSX. Veja a mensagem acima.")

        out.write_bytes(r.content)
        logging.info("Arquivo salvo em: %s", out.resolve())
        return out

def calcula_promedio(arquivo: str):
    logging.info("Calculando promedio para %s", arquivo)
    Path("promedios").mkdir(parents=True, exist_ok=True)
    caminho = Path("produtos") / arquivo

    df = pd.read_excel(caminho, header=2)

    df = df.iloc[:, :5].copy()
    df.columns = ["Año", "Mes", "Fecha", "Average of PRECIO", "PRODUCTO"]

    # Converter Data do Fecha
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")

    def _parse_preco(valor):
        if pd.isna(valor):
            return None
        if isinstance(valor, (int, float)):
            return float(valor)

        s = str(valor).strip().replace("$", "").replace(" ", "")
        if not s:
            return None

        # tratando caso o valor esteja em real $
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        # Se só vírgula, assume decimal com vírgula
        elif "," in s:
            s = s.replace(",", ".")
        # Se só ponto, mantém como decimal

        try:
            return float(s)
        except ValueError:
            return None

    df["PRECIO_num"] = df["Average of PRECIO"].apply(_parse_preco)

    #Criar mês em texto
    df["ano"] = df["Mes"].astype(str).str.slice(0, 4).astype("Int64")
    df["mes_num"] = df["Mes"].astype(str).str.slice(5, 7).astype("Int64")

    mes_pt = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    df["Mês"] = df["mes_num"].map(mes_pt)


    #Agrupar por ano + mês + produto e calculo da média mensal
    df_valid = df.dropna(subset=["ano", "mes_num", "PRODUCTO", "PRECIO_num"]).copy()

    tabela_mensal = (
        df_valid.groupby(["ano", "mes_num", "Mês", "PRODUCTO"], as_index=False)["PRECIO_num"]
        .mean()
        .rename(columns={"ano": "Año", "PRECIO_num": "Promedio año actual"})
        .sort_values(["Año", "mes_num", "PRODUCTO"])
    )
    #Formatando a saída
    tabela_mensal["Promedio año actual"] = tabela_mensal["Promedio año actual"].round(0).astype("Int64")
    tabela_mensal["Promedio año actual"] = tabela_mensal["Promedio año actual"].map(
        lambda x: f"$ {x:,}".replace(",", ".") if pd.notna(x) else None
    )

    tabela_mensal = tabela_mensal[["Año", "Mês", "Promedio año actual", "PRODUCTO"]]
    tabela_mensal.to_excel(f"promedios/{arquivo}", index=False)
    logging.info("Promedio salvo em promedios/%s", arquivo)

def corrige_filename(name: str) -> str:
    return re.sub(r'[\/\\:\*\?"<>\|]', "_", name)

if __name__ == "__main__":
    
    produtos = []
    with open("produtos.txt", "r", encoding="utf-8") as f:
        for line in f:
            produto = line.strip()
            if produto:
                produtos.append(produto)
    logging.info("Realizando o download dos dados de cada produto...")
    for p in produtos:
        logging.info("Produto: %s", p)
        atualizar_payload("payload.json", p)
        nome_arquivo = corrige_filename(f"{p}.xlsx")
        path = download_xlsx(f"produtos/{nome_arquivo}")
        
    logging.info("Iniciando o cálculo dos promedios para cada produto...")
    pasta_produtos = Path("produtos")
    for arquivo in pasta_produtos.iterdir():
        if arquivo.is_file() and arquivo.suffix.lower() in {".xlsx"}:
            logging.info("Processando arquivo: %s", arquivo.name)
            calcula_promedio(arquivo.name)