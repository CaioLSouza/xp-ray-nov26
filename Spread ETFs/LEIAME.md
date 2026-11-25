# Barras intraday de BID/ASK de ETFs (LSEG Workspace)

Baixa as barras intraday dos ETFs listados na planilha e entrega um CSV bruto no formato pedido pelo Research (versão de 2026-11-25), sem filtro nem cálculo. Os campos são compra, venda, quantidade no topo do livro e último preço. À parte, gera um Excel com o spread e a profundidade já calculados.

## Arquivos
- `inputs_spread_etfs.xlsx`: a planilha que você edita (células amarelas).
  - aba **ETFs**: um ticker por linha, sem o `.SA` (hoje são 21).
    - O RIC padrão é `<ticker>.SA`; preencha a coluna `ric` só se for diferente.
    - `ativo = N` tira o ETF da rodada sem apagar a linha.
  - aba **Parametros**:
    - período;
    - tamanho da barra (5 min);
    - janela do dia (10h15 a 16h45);
    - nome dos campos da API;
    - opções.
- `spread_etfs.py`: o script.
- `rodar.bat`: roda o período da planilha (primeira carga).
- `rodar_mes_anterior.bat`: roda só o mês anterior (rotina mensal).

## Como rodar (no PC com o Workspace)
1. Uma vez só: `pip install lseg-data pandas openpyxl`.
2. Gere uma app key no Workspace (busca: "App Key Generator") e guarde numa variável de ambiente, num terminal seu: `setx LSEG_APP_KEY "sua-chave"`. Depois feche e abra o terminal.
3. Abra e logue o Workspace.
4. Rode o script:
   - **Primeira carga** (período da planilha, hoje de 25/11/2025 a 24/11/2026): `python spread_etfs.py` ou `rodar.bat`.
   - **Todo mês** (só o mês calendário anterior): `python spread_etfs.py --mes-anterior` ou `rodar_mes_anterior.bat`. O mesmo vale com `periodo = mes_anterior` na planilha.
     - Em dezembro, por exemplo, ele baixa de 01/11 a 30/11. Os dias que já vieram na carga voltam de novo, e vale a versão nova.
     - Pode ir para o Agendador de Tarefas do Windows, desde que o Workspace esteja aberto e logado na hora.
   - Para usar outra planilha: `python spread_etfs.py outra.xlsx` (combina com `--mes-anterior`).

## O que sai em `saida\`
**`etfs_bid_ask_5min_AAAAMMDD_AAAAMMDD.csv`**
- Uma linha por ETF a cada 5 minutos.
- UTF-8, vírgula como separador, ponto decimal e sem coluna de índice.
- As colunas são estas, nesta ordem:

| coluna | exemplo | campo |
|---|---|---|
| data_hora | 2026-11-18 10:15:00 | horário de Brasília |
| ticker | BOVA11 | sem o .SA |
| compra | 173.50 | BID |
| venda | 173.52 | ASK |
| qtd_compra | 5000 | BIDSIZE (cotas, inteiro) |
| qtd_venda | 3200 | ASKSIZE (cotas, inteiro) |
| ultimo | 173.51 | TRDPRC_1 (último preço) |

Cada período tem o seu nome de arquivo, então uma rodada mensal não sobrescreve a carga nem a do mês anterior. Rodar de novo o mesmo período substitui os arquivos dele.

**`spread_<...>.xlsx`** (com `gerar_spread = S`): o spread e a profundidade já calculados, num Excel formatado. O CSV bruto não muda.

| Aba | Conteúdo |
|---|---|
| Resumo | uma linha por ETF: pregões, spread médio, média dos últimos 21 pregões, mediana, P10 e P90 dos dias, spread em R$, profundidade média no topo (cotas e R$ mil), % das barras com cotação válida |
| Mensal (bps) | spread médio por mês × ETF |
| Diário (bps) | spread médio por pregão × ETF |
| Perfil intradiário (bps) | spread médio por horário (10:15 … 16:45) × ETF, no período todo |
| Diário detalhado | por ETF e pregão: barras, barras válidas, médio, mediana, P90, mínimo, máximo, R$ e profundidade |
| Notas | método |

- **Conta:** spread da barra = (venda − compra) / ponto médio × 10.000, em bps. Entram só as barras válidas (compra > 0, venda > 0 e venda > compra).
- **Profundidade no topo:** média de (qtd_compra + qtd_venda) / 2 nas barras válidas. Em R$ mil, cada lado vezes o seu preço.
- **Dia:** o spread do dia é a média das barras válidas. Pregão com menos de `min_barras_validas` barras válidas fica sem spread.
- **Período:** o resumo e o mensal são a média simples dos dias.
- **Refazer só o Excel**, sem chamar a API: `python spread_etfs.py saida\<arquivo bruto>.csv`.

**`relatorio_<...>.txt`**: o que vai junto com cada arquivo.
- A data de execução e o período.
- **Falhas:**
  - ETF com erro na chamada, que fica fora do arquivo, enquanto os outros seguem;
  - ETF sem nenhum dado (conferir o RIC);
  - campo que veio vazio no período todo (conferir o nome no DIB);
  - pregão faltando num ETF quando outros ETFs tiveram dado no dia;
  - dias úteis sem dado em nenhum ETF (em geral, feriado da B3).
- Os nomes exatos dos campos, da função e do intervalo usados na chamada.
- Desde quando existe histórico intradiário, por ETF. Na primeira carga, se o período já começa com dado, o script volta mês a mês para descobrir a data. Na rotina mensal ele não sonda.
- Linhas por ETF e o log da rodada.

## Como a chamada é feita
- `get_history(universe=<RIC>, fields=["BID", "ASK", "BIDSIZE", "ASKSIZE", "TRDPRC_1"], interval="5min", start, end)`. Cada barra traz o último valor do intervalo.
- Os nomes dos campos ficam na aba Parametros (`campo_compra`, `campo_venda`, `campo_qtd_compra`, `campo_qtd_venda` e `campo_ultimo`). Se o DIB mostrar outro nome, basta trocar ali.
- Compra e venda são obrigatórios: se a API não devolver esses campos, o ETF dá erro. Se os outros vierem vazios, o relatório avisa.
- Vai um mês calendário por chamada. Se a resposta vier cortada pelo limite de linhas, o mês é dividido sozinho.
- Os meses fechados ficam em `cache\`. Rodar de novo só baixa o que falta (`usar_cache = S`).
  - O cache é separado por lista de campos: ao trocar um campo, aquele mês é baixado de novo.
- A API devolve em UTC, e o CSV sai no horário de Brasília.

## Tratamento (só o pedido)
- **Janela:** 10:15 a 16:45, inclusive, uma linha a cada 5 minutos para cada ETF, só nos pregões com dado.
- **Barra sem atualização:** repete o último valor do mesmo dia, campo a campo, nunca do dia anterior.
  - As barras do próprio dia entre 10:00 e 10:15 servem de ponto de partida; as do leilão de abertura, não.
  - Se o dia ainda não tinha oferta até aquele horário, a linha fica vazia, nunca zero.
- **Filtros:** nenhum. Livro cruzado, oferta zerada e valores estranhos vão como vieram.

## Pontos a confirmar na primeira rodada
- **Campos de quantidade e último preço:** se `BIDSIZE`, `ASKSIZE` ou `TRDPRC_1` vierem vazios, o relatório avisa. Confira o nome no Data Item Browser (DIB) e troque na aba Parametros.
- **Carimbo de hora:** ver se o da barra marca o início ou o fim dos 5 minutos (documentação da LSEG).
- **Campo diário de spread médio:** se existe um pronto, só como referência. O script não verifica; buscar "spread" no DIB e ler a definição.

## Se der erro
- **"A resposta não trouxe ['BID', 'ASK']"**: a mensagem lista as colunas que vieram.
- **ETF com 0 barras em todos os meses**: conferir o RIC no Workspace (ex.: `BOVA11.SA`).
- **Erro de permissão ao gravar**: o CSV ou o Excel está aberto no Excel; feche e rode de novo, porque o cache já baixado fica.
