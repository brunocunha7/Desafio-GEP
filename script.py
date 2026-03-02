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

def periodo_mes(ano: int, mes: int):

    # primeiro dia do mês
    data_inicio = date(ano, mes, 1)
    # para calcular o último dia, avançamos para o primeiro dia do próximo mês
    if mes == 12:
        proximo = date(ano + 1, 1, 1)
    else:
        proximo = date(ano, mes + 1, 1)
    data_fim = proximo - timedelta(days=1)
    return data_inicio, data_fim

def atualizar_payload(caminho_json: str, novo_produto: str, data_inicio: date, data_fim: date, departamento: str = "Nacional") -> None:
    logging.info("Atualizando payload para produto: %s", novo_produto)
    with open(caminho_json, "r", encoding="utf-8") as f:
        payload = json.load(f)

    comandos = payload["executeSemanticQueryRequest"]["queries"][0]["Query"]["Commands"]
    where_list = comandos[0]["SemanticQueryDataShapeCommand"]["Query"]["Where"]

    # Atualiza intervalos de data no primeiro filtro (condição And)
    and_cond = where_list[0]["Condition"]["And"]
    # data_inicio -> Left comparison right literal
    and_cond["Left"]["Comparison"]["Right"]["Literal"]["Value"] = (
        f"datetime'{data_inicio.isoformat()}T00:00:00'"
    )
    # data_fim -> Right comparison right literal
    and_cond["Right"]["Comparison"]["Right"]["Literal"]["Value"] = (
        f"datetime'{data_fim.isoformat()}T00:00:00'"
    )

    # percorre os filtros restantes e ajusta conforme o tipo
    for clause in where_list[1:]:
        cond = clause.get("Condition", {})
        if "In" not in cond:
            continue
        exprs = cond["In"].get("Expressions", [])
        if not exprs:
            continue
        prop = exprs[0].get("Column", {}).get("Property")
        if prop == "DEPARTAMENTO":
            # sempre Nacional
            cond["In"]["Values"][0][0]["Literal"]["Value"] = f"'{departamento}'"
        elif prop == "PRODUCTO":
            cond["In"]["Values"][0][0]["Literal"]["Value"] = f"'{novo_produto}'"

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

def calcula_promedio(arquivo: str, produto: str) -> pd.DataFrame:
    logging.info("Calculando promedio para %s", arquivo)
    caminho = Path("produtos") / produto / arquivo

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

    # extrair ano e mês numérico para agrupar e construir data
    df["ano"] = df["Mes"].astype(str).str.slice(0, 4).astype("Int64")
    df["mes_num"] = df["Mes"].astype(str).str.slice(5, 7).astype("Int64")

    #Agrupar por ano + mês + produto e calculo da média mensal
    df_valid = df.dropna(subset=["ano", "mes_num", "PRODUCTO", "PRECIO_num"]).copy()

    tabela_mensal = (
        df_valid.groupby(["ano", "mes_num", "PRODUCTO"], as_index=False)["PRECIO_num"]
        .mean()
        .rename(columns={"ano": "Año", "mes_num": "mes_num", "PRECIO_num": "valor"})
        .sort_values(["Año", "mes_num", "PRODUCTO"])
    )

    # preparar resultado no formato solicitado
    # Data: primeiro dia do mês como string dd/mm/YYYY
    tabela_mensal["Data"] = tabela_mensal.apply(
        lambda row: f"01/{int(row['mes_num']):02d}/{int(row['Año'])}", axis=1
    )
    tabela_mensal["Referencia"] = tabela_mensal["PRODUCTO"]

    # valor deve ser inteiro sem formatação
    tabela_mensal["valor"] = tabela_mensal["valor"].round(0).astype("Int64")

    resultado = tabela_mensal[["Referencia", "Data", "valor"]].copy()
    return resultado

def corrige_filename(name: str) -> str:
    return re.sub(r'[\/\\:\*\?"<>\|]', "_", name)



if __name__ == "__main__":
    
    produtos = ["Azúcar Blanco", "Maíz Amarillo Nacional"]
    ano = 2025
    resultados = []  # coletar dataframes de cada execução
     
    for p in produtos:
        for mes in range(1, 13):
            data_inicio, data_fim = periodo_mes(ano, mes)
            logging.info("Produto: %s", p)
            atualizar_payload("payload.json", p, data_inicio, data_fim, departamento="Nacional")
            nome_arquivo = corrige_filename(f"{mes}-{p}.xlsx")
            Path(f"produtos/{p}").mkdir(parents=True, exist_ok=True)
            path = download_xlsx(f"produtos/{p}/{nome_arquivo}")

            # chama a função para calcular o promedio
            df_prom = calcula_promedio(nome_arquivo, p)
            logging.info("Promedio calculado para %s:\n%s", nome_arquivo, df_prom)
            resultados.append(df_prom)

    # combinar e salvar CSV se houver resultados
    if resultados:
        combinado = pd.concat(resultados, ignore_index=True)
        combinado.to_csv("promedios.csv", index=False)
        logging.info("Arquivo CSV combinado salvo como promedios.csv")



