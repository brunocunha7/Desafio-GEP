# Desafio Técnico: Engenharia Reversa de API em Dashboards (Power BI)

Este repositório contém o  script Python (`script.py`) que:

* baixa relatórios XLSX do dashboard Power BI (Bolsa Mercantil da
  Colômbia),
* processa os dados para extrair médias mensais de preços por produto, e
* combina os resultados em um arquivo CSV final (`promedios.csv`).

## Etapas executadas pelo script

1. **Inicialização dos objetos principais**
   * `PowerBIReport` é instanciado com as URLs/IDs necessários e o caminho do
     payload JSON.
   * `DataProcessor` é criado para cálculos de datas e manipulação de dados.

2. **Iteração sobre produtos e meses**
   * Lista de produtos, passadas no documento do desafio ("Azúcar Blanco", "Maíz Amarillo Nacional").
   * Para cada produto, percorre-se os 12 meses de um ano fixo (2025).

3. **Atualização do payload de pesquisa**
   * O JSON (`payload.json`) usado para exportação é modificado com o
     produto e o intervalo de datas correspondente ao mês atual. Foi colocado como um arquivo externo para não poluir o código e ficar melhor a visualização.

4. **Download do relatório**
   * Chamada POST para a API de exportação do Power BI usando token obtido via
     requisição GET.
   * Salva o XLSX em `produtos/<produto>/<mês>-<produto>.xlsx`.

5. **Cálculo do promedio (média mensal)**
   * Lê o XLSX baixado com pandas, ajusta colunas e converte preços para
     numérico.
   * Agrupa por ano, mês e produto para calcular a média de preço.
   * Gera um DataFrame com colunas `Referencia`, `Data` (primeiro dia do mês)
     e `valor` (média arredondada).
   * O DataFrame é acumulado em uma lista de resultados.

6. **Combinação e exportação final**
   * Todos os DataFrames gerados na iteração são concatenados.
   * O resultado combinado é escrito em `promedios.csv`.

## Execução

1. Ative o ambiente virtual Python:
   ```bash
   source venv/bin/activate
   ```
2. Assegure-se de que `payload.json` esteja presente.
3. Execute:
   ```bash
   python script.py
   ```

Os arquivos XLSX serão baixados para a pasta `produtos/` organizada por
produto, e o CSV final aparecerá na raiz. As pastas são criadas automaticamente.

## Observações

* O script usa timeouts e registraprogresso via `logging`.
* Interrupções manuais (`Ctrl+C`) são registradas e cortam a execução atual.

---

Esta documentação descreve de forma clara o que o código faz e as etapas
realizadas durante sua execução.

Desenvolvido por `Bruno da Cunha Castro`