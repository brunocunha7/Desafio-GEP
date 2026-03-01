# Script de Download e Cálculo de Promedios

Este repositório contém um script (`script.py`) que:
- gera/atualiza um `payload.json` com filtros por produto e data;
- faz requisições ao endpoint do Power BI para exportar dados por produto em XLSX;
- calcula promedios mensais a partir dos arquivos baixados e salva em `promedios/`.

Pré-requisitos
- Python 3.8+ (foi desenvolvido em 3.14).
- Repositório com os arquivos:
  - `script.py` (o script principal)
  - `payload.json` (modelo de payload usado pela API)
  - `produtos.txt` (lista de produtos, um por linha)

Instalação (recomendado usar virtualenv)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Se você não tiver um `requirements.txt`, instale diretamente:

```bash
pip install requests pandas openpyxl
```

Uso

1. Coloque os produtos que quer processar em `produtos.txt` (um por linha).
2. Não é necessário ajustar o arquivo `payload.json`, o script atualiza apenas os filtros `PRODUCTO` e a data final para o dia de ontem. Dependendo do horário de atualização do site, o script pode ser modificado para utilizar o dia de hoje.
3. Execute:

```bash
python3 script.py
```

O script fará:
- criar/usar a pasta `produtos/` para salvar os arquivos XLSX baixados;
- criar/usar a pasta `promedios/` para salvar os arquivos resultantes de média mensal (XLSX);
- registrar eventos no console via `logging` (nível `INFO`).

Arquivos gerados
- `produtos/<produto>.xlsx` — arquivos originais baixados do Power BI.
- `promedios/<produto>.xlsx` — tabelas com o promedio mensal calculado.

Script desenvolvido por `Bruno da Cunha Castro`