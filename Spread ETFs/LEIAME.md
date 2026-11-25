# Barras intraday de BID/ASK de ETFs (LSEG Workspace)

Baixa as barras intraday de BID, ASK e, opcionalmente, do último preço dos ETFs listados na planilha, e entrega um CSV bruto no formato pedido pelo Research (2026-11-25), sem filtro nem cálculo. À parte, gera um Excel com o spread já calculado.

## Arquivos
- `inputs_spread_etfs.xlsx`: a planilha que você edita (células amarelas).
  - aba **ETFs**: um ticker por linha, sem o `.SA`.
    - O RIC padrão é `<ticker>.SA`; preencha a coluna `ric` só se for diferente.
    - `ativo = N` tira o ETF da rodada sem apagar a linha.
  - aba **Parametros**: período, tamanho da barra (5 min), janela do dia (10h15 a 16h45) e opções.
- `spread_etfs.py`: o script.
- `rodar.bat`: atalho para rodar com dois cliques.

## Como rodar (no PC com o Workspace)
1. Uma vez só: `pip install lseg-data pandas openpyxl`.
2. Gere uma app key no Workspace (busca: "App Key Generator") e guarde numa variável de ambiente, num terminal seu: `setx LSEG_APP_KEY "sua-chave"`. Depois feche e abra o terminal.
3. Abra e logue o Workspace.
4. Rode `python spread_etfs.py` (ou `rodar.bat`). Para usar outra planilha: `python spread_etfs.py outra.xlsx`.

## O que sai em `saida\`
**`etfs_bid_ask_5min_<inicio>_<fim>.csv`**: uma linha por ETF a cada 5 minutos, com estas colunas, nesta ordem.

| coluna | exemplo | observação |
|---|---|---|
| data_hora | 2026-11-18 10:15:00 | horário de Brasília |
| ticker | BOVA11 | sem o .SA |
| compra | 173.50 | BID, ponto como separador decimal |
| venda | 173.60 | ASK |
| ultimo | 173.55 | último preço (TRDPRC_1); sai se `incluir_ultimo = N` |

**`spread_<...>.xlsx`** (com `gerar_spread = S`): o spread já calculado, num Excel formatado. O CSV bruto não muda.

| Aba | Conteúdo |
|---|---|
| Resumo | uma linha por ETF: pregões, spread médio, média dos últimos 21 pregões, mediana, P10 e P90 dos dias, spread em R$, % das barras com cotação válida |
| Mensal (bps) | spread médio por mês × ETF |
| Diário (bps) | spread médio por pregão × ETF |
| Perfil intradiário (bps) | spread médio por horário (10:15 … 16:45) × ETF, no período todo |
| Diário detalhado | por ETF e pregão: barras, barras válidas, médio, mediana, P90, mínimo, máximo e R$ |
| Notas | método |

- **Conta:** spread da barra = (venda − compra) / ponto médio × 10.000, em bps. Entram só as barras válidas (compra > 0, venda > 0 e venda > compra).
- **Dia:** o spread do dia é a média das barras válidas. Pregão com menos de `min_barras_validas` barras válidas fica sem spread.
- **Período:** o resumo e o mensal são a média simples dos dias.
- **Refazer só o Excel**, sem chamar a API: `python spread_etfs.py saida\<arquivo bruto>.csv`.

**`relatorio_<...>.txt`**: o que ela pediu para mandar junto.
- Desde quando existe histórico intradiário, por ETF. Se o período pedido já começa com dado, o script volta mês a mês para descobrir a data.
- Os nomes exatos da função, dos campos e do intervalo usados na chamada.
- Linhas por ETF, linhas sem compra ou venda e o log da rodada.

## Como a chamada é feita
- `get_history(universe=<RIC>, fields=["BID", "ASK", "TRDPRC_1"], interval="5min", start, end)`. Cada barra traz o último valor do intervalo.
- Vai um mês calendário por chamada. Se a resposta vier cortada pelo limite de linhas, o mês é dividido sozinho.
- Os meses fechados ficam em `cache\`. Rodar de novo só baixa o que falta (`usar_cache = S`).
- A API devolve em UTC, e o CSV sai no horário de Brasília.

## Tratamento (só o pedido)
- **Janela:** 10:15 a 16:45, inclusive, uma linha a cada 5 minutos, só nos pregões com dado.
- **Barra que falta:** repete o último valor do mesmo dia, nunca do dia anterior. As barras do próprio dia entre 10:00 e 10:15 servem de ponto de partida; as do leilão de abertura, não. Se o dia ainda não tinha cotação até aquele horário, a linha sai com compra e venda vazios.
- **Filtros:** nenhum. Livro cruzado, oferta zerada e valores estranhos vão como vieram.

## Pontos a confirmar na primeira rodada
- **Campo do último preço:** se o `TRDPRC_1` vier vazio, o relatório avisa; conferir o nome no Data Item Browser (DIB).
- **Carimbo de hora:** ver se o da barra marca o início ou o fim dos 5 minutos (documentação da LSEG).
- **Campo diário de spread médio:** se existe um pronto, só como referência. O script não verifica; buscar "spread" no DIB e ler a definição.

## Se der erro
- **"A resposta não trouxe ['BID', 'ASK']"**: a mensagem lista as colunas que vieram.
- **ETF com 0 barras em todos os meses**: conferir o RIC no Workspace (ex.: `BOVA11.SA`).
