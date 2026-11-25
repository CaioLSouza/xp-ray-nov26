"""
Barras intraday de BID e ASK (e, opcionalmente, último preço) de ETFs da B3, pela LSEG (Refinitiv Workspace),
no formato pedido pelo Research (2026-11-25). Gera também um Excel à parte com o spread calculado.

Uso (no PC com o Workspace aberto e logado):
    python spread_etfs.py                       # lê inputs_spread_etfs.xlsx na mesma pasta
    python spread_etfs.py outra_planilha.xlsx   # ou outra planilha de entrada

O que faz:
  1. Lê os ETFs e os parâmetros na planilha de entrada.
  2. Para cada ETF, baixa as barras (padrão: 5 minutos) de BID, ASK e TRDPRC_1, um mês por chamada.
     Usa o último valor de cada barra, como a API devolve. Se a resposta vier cortada pelo limite de
     linhas, divide o mês sozinho. Os meses já fechados ficam em cache.
  3. Converte o horário de UTC para Brasília e monta, para cada pregão, uma linha a cada 5 minutos de
     10:15 a 16:45 (inclusive), sem buracos. Barra que falta repete o último valor do MESMO dia (nunca do
     dia anterior). Não filtra mais nada: livro cruzado, oferta zerada etc. vão como vieram.
  4. Grava saida\\<arquivo>.csv (data_hora, ticker, compra, venda, ultimo) e um relatorio_*.txt com o que
     foi pedido para mandar junto: desde quando há histórico, campos e intervalo usados na chamada.

Requer: pip install lseg-data pandas openpyxl   (ou a biblioteca antiga refinitiv-data)
App key: variável de ambiente LSEG_APP_KEY (gerada no App Key Generator do Workspace),
ou o arquivo de configuração padrão da biblioteca. Não grave a chave na planilha.
"""
from __future__ import annotations

import os
import sys
import time
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

PASTA = Path(__file__).resolve().parent
TZ = "America/Sao_Paulo"
CAMPO_BID, CAMPO_ASK, CAMPO_ULTIMO = "BID", "ASK", "TRDPRC_1"
LOG: list[str] = []


def log(msg: str) -> None:
    linha = f"{dt.datetime.now():%H:%M:%S}  {msg}"
    print(linha, flush=True)
    LOG.append(linha)


# --------------------------------------------------------------------------------------------
# Planilha de entrada
# --------------------------------------------------------------------------------------------
def _sim(v) -> bool:
    return str(v).strip().upper() in ("S", "SIM", "Y", "YES", "1", "TRUE", "X")


def _hora(v) -> dt.time:
    if isinstance(v, dt.time):
        return v
    if isinstance(v, dt.datetime):
        return v.time()
    if isinstance(v, (float, int)) and 0 <= v < 1:  # fração de dia do Excel
        seg = int(round(v * 86400))
        return dt.time(seg // 3600, (seg % 3600) // 60)
    return dt.datetime.strptime(str(v).strip()[:5], "%H:%M").time()


def ler_entrada(caminho: Path):
    xls = pd.ExcelFile(caminho)

    etfs = pd.read_excel(xls, "ETFs", dtype=str).fillna("")
    etfs.columns = [c.strip().lower() for c in etfs.columns]
    for col in ("ric", "ativo"):
        if col not in etfs.columns:
            etfs[col] = ""
    etfs["ticker"] = etfs["ticker"].str.strip().str.upper().str.replace(".SA", "", regex=False)
    etfs = etfs[etfs["ticker"] != ""]
    etfs = etfs[etfs["ativo"].map(_sim) | (etfs["ativo"].str.strip() == "")].copy()
    etfs["ric"] = [r.strip() if r.strip() else f"{t}.SA" for t, r in zip(etfs["ticker"], etfs["ric"])]
    etfs = etfs.drop_duplicates("ticker").reset_index(drop=True)
    if etfs.empty:
        raise SystemExit("Nenhum ETF ativo na aba ETFs.")

    par = pd.read_excel(xls, "Parametros")
    par.columns = [c.strip().lower() for c in par.columns]
    P = {str(k).strip(): v for k, v in zip(par["parametro"], par["valor"]) if str(k).strip() and str(k) != "nan"}

    def p(nome, padrao):
        v = P.get(nome)
        return padrao if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "" else v

    ontem = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)
    cfg = {
        "data_inicio": pd.Timestamp(p("data_inicio", ontem - pd.Timedelta(days=365))).normalize(),
        "data_fim": pd.Timestamp(p("data_fim", ontem)).normalize(),
        "intervalo_min": int(p("intervalo_min", 5)),
        "janela_inicio": _hora(p("janela_inicio", "10:15")),
        "janela_fim": _hora(p("janela_fim", "16:45")),
        "abertura_continuo": _hora(p("abertura_continuo", "10:00")),
        "incluir_ultimo": _sim(p("incluir_ultimo", "S")),
        "gerar_spread": _sim(p("gerar_spread", "S")),
        "min_barras_validas": int(p("min_barras_validas", 20)),
        "sondar_historico": _sim(p("sondar_historico", "S")),
        "pausa_seg": float(p("pausa_seg", 0.5)),
        "usar_cache": _sim(p("usar_cache", "S")),
        "pasta_saida": str(p("pasta_saida", "saida")),
        "arquivo_saida": str(p("arquivo_saida", "")).strip(),
    }
    if cfg["data_fim"] < cfg["data_inicio"]:
        raise SystemExit("data_fim antes de data_inicio na aba Parametros.")
    if not cfg["abertura_continuo"] <= cfg["janela_inicio"] < cfg["janela_fim"]:
        raise SystemExit("Na aba Parametros, a ordem tem de ser abertura_continuo <= janela_inicio < janela_fim.")
    cfg["intervalo_api"] = f"{cfg['intervalo_min']}min"
    return etfs, cfg


# --------------------------------------------------------------------------------------------
# API da LSEG
# --------------------------------------------------------------------------------------------
def importar_api():
    try:
        import lseg.data as ld  # biblioteca atual
    except ImportError:
        try:
            import refinitiv.data as ld  # biblioteca antiga, mesma interface
        except ImportError:
            raise SystemExit("Instale a biblioteca da LSEG: pip install lseg-data")
    return ld


def abrir_sessao(ld):
    chave = os.environ.get("LSEG_APP_KEY") or os.environ.get("REFINITIV_APP_KEY")
    if chave:
        ld.open_session(app_key=chave)
    else:
        ld.open_session()
    log("Sessão LSEG aberta.")


def _vazio() -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype=float) for c in (CAMPO_BID, CAMPO_ASK, CAMPO_ULTIMO)},
                        index=pd.DatetimeIndex([], tz="UTC"))


def _normalizar(df: pd.DataFrame | None) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return _vazio()
    df = df.copy()
    if isinstance(df.columns, pd.MultiIndex):  # alguns retornos vêm com (RIC, campo)
        df.columns = [c[-1] for c in df.columns]
    df.columns = [str(c).upper() for c in df.columns]
    faltando = [c for c in (CAMPO_BID, CAMPO_ASK) if c not in df.columns]
    if faltando:
        raise RuntimeError(f"A resposta não trouxe {faltando}. Colunas recebidas: {list(df.columns)}")
    if CAMPO_ULTIMO not in df.columns:
        df[CAMPO_ULTIMO] = np.nan
    df = df[[CAMPO_BID, CAMPO_ASK, CAMPO_ULTIMO]].apply(pd.to_numeric, errors="coerce")
    idx = pd.to_datetime(df.index)
    idx = idx.tz_localize("UTC") if idx.tz is None else idx.tz_convert("UTC")  # a LSEG devolve em UTC
    df.index = idx
    return df[~df.index.isna()].sort_index()


def chamar_api(ld, ric, campos, intervalo, ini_utc, fim_utc, pausa) -> pd.DataFrame:
    for tentativa in range(1, 4):
        try:
            df = ld.get_history(
                universe=ric, fields=campos, interval=intervalo,
                start=ini_utc.strftime("%Y-%m-%dT%H:%M:%S"), end=fim_utc.strftime("%Y-%m-%dT%H:%M:%S"),
            )
            time.sleep(pausa)
            return _normalizar(df)
        except RuntimeError:
            raise
        except Exception as e:  # noqa: BLE001
            texto = str(e).lower()
            if "no data" in texto or "not found" in texto or "empty" in texto:
                time.sleep(pausa)
                return _vazio()
            espera = 5 * tentativa
            log(f"   erro na API ({ric}, {ini_utc:%Y-%m-%d}): {e}. Nova tentativa em {espera}s.")
            time.sleep(espera)
    raise RuntimeError(f"Falhou 3 vezes: {ric} {ini_utc} → {fim_utc}")


def _parece_cortado(df, ini_utc, fim_utc) -> bool:
    """A resposta parece cortada pelo limite de linhas da API?

    Sinais: falta dia útil (seg-sex) da janela antes do 1º ou depois do último dia que veio, ou o 1º/último dia
    veio com bem menos barras que o dia típico da resposta (o corte caiu no meio de um pregão). Um falso alarme
    (feriado, meio pregão) só custa algumas chamadas a mais.
    """
    dias = pd.date_range(ini_utc.tz_convert(TZ).normalize(),
                         (fim_utc - pd.Timedelta(seconds=1)).tz_convert(TZ).normalize(), freq="D")
    dias = {d.date() for d in dias if d.dayofweek < 5}
    por_dia = pd.Series(df.index.tz_convert(TZ).date).value_counts().sort_index()
    d0, d1 = por_dia.index[0], por_dia.index[-1]
    if any(d < d0 or d > d1 for d in dias):
        return True
    tipico = por_dia.median()
    return por_dia.iloc[0] < 0.5 * tipico or por_dia.iloc[-1] < 0.5 * tipico


def _meio_na_virada_do_dia(ini_utc, fim_utc):
    """Ponto de corte na meia-noite (Brasília) mais próxima do meio, para não partir um pregão em dois."""
    meio = (ini_utc + (fim_utc - ini_utc) / 2).tz_convert(TZ).normalize().tz_convert("UTC")
    if not ini_utc < meio < fim_utc:
        meio = (ini_utc + (fim_utc - ini_utc) / 2).tz_convert(TZ).normalize() + pd.Timedelta(days=1)
        meio = meio.tz_convert("UTC")
    return meio if ini_utc < meio < fim_utc else None


def baixar(ld, ric, campos, intervalo, ini_utc, fim_utc, pausa) -> pd.DataFrame:
    """Baixa [ini, fim); se a resposta parecer cortada pelo limite de linhas, divide a janela na virada do dia."""
    df = chamar_api(ld, ric, campos, intervalo, ini_utc, fim_utc, pausa)
    meio = _meio_na_virada_do_dia(ini_utc, fim_utc) if len(df) > 500 else None
    if meio is not None and _parece_cortado(df, ini_utc, fim_utc):
        df = pd.concat([baixar(ld, ric, campos, intervalo, ini_utc, meio, pausa),
                        baixar(ld, ric, campos, intervalo, meio, fim_utc, pausa)])
    df = df[(df.index >= ini_utc) & (df.index < fim_utc)]
    return df[~df.index.duplicated(keep="last")]


def meses(ini: pd.Timestamp, fim_incl: pd.Timestamp):
    """Fatias [a, b) de mês calendário cobrindo ini..fim_incl."""
    a = ini
    fim = fim_incl + pd.Timedelta(days=1)
    while a < fim:
        b = min((a + pd.offsets.MonthBegin(1)).normalize(), fim)
        yield a, b
        a = b


def campos_da_chamada(cfg):
    return [CAMPO_BID, CAMPO_ASK] + ([CAMPO_ULTIMO] if cfg["incluir_ultimo"] else [])


def baixar_periodo(ld, ric, cfg) -> pd.DataFrame:
    cache_dir = PASTA / "cache"
    cache_dir.mkdir(exist_ok=True)
    hoje = pd.Timestamp.today().normalize()
    campos = campos_da_chamada(cfg)
    partes = []
    for a, b in meses(cfg["data_inicio"], cfg["data_fim"]):
        arq = cache_dir / f"{ric.replace('.', '_')}_{cfg['intervalo_api']}_{'+'.join(campos)}_{a:%Y%m%d}_{b:%Y%m%d}.csv"
        if cfg["usar_cache"] and arq.exists() and b <= hoje:
            df = pd.read_csv(arq, index_col=0)
            df.index = pd.to_datetime(df.index, utc=True)
        else:
            df = baixar(ld, ric, campos, cfg["intervalo_api"],
                        a.tz_localize(TZ).tz_convert("UTC"), b.tz_localize(TZ).tz_convert("UTC"), cfg["pausa_seg"])
            if b <= hoje:  # só guarda mês fechado
                df.to_csv(arq)
        log(f"   {ric} {a:%d/%m/%Y} a {b - pd.Timedelta(days=1):%d/%m/%Y}: {len(df)} barras")
        partes.append(df)
    out = pd.concat(partes) if partes else _vazio()
    return out[~out.index.duplicated(keep="last")].sort_index()


def sondar_inicio(ld, ric, cfg, primeira_no_periodo):
    """Se o período pedido já começa com dado, volta mês a mês para achar desde quando existe histórico."""
    if primeira_no_periodo is None:
        return None, "sem dado no período pedido"
    if primeira_no_periodo.tz_convert(TZ).date() > (cfg["data_inicio"] + pd.Timedelta(days=7)).date():
        return primeira_no_periodo, "o histórico começa dentro do período pedido"
    if not cfg["sondar_historico"]:
        return None, "existe pelo menos desde o início do período (sondagem desligada)"
    campos = campos_da_chamada(cfg)
    primeira = primeira_no_periodo
    fim = cfg["data_inicio"]
    for _ in range(36):  # no máximo 3 anos para trás
        ini = (fim - pd.offsets.MonthBegin(1)).normalize()
        df = chamar_api(ld, ric, campos, cfg["intervalo_api"],
                        ini.tz_localize(TZ).tz_convert("UTC"), fim.tz_localize(TZ).tz_convert("UTC"), cfg["pausa_seg"])
        if df.empty:
            return primeira, "primeira barra devolvida pela API (sondagem mês a mês para trás)"
        primeira = df.index.min()
        fim = ini
    return primeira, "a API devolveu dado nos 36 meses sondados; o histórico vai além disso"


# --------------------------------------------------------------------------------------------
# Tratamento
# --------------------------------------------------------------------------------------------
def montar_grade(df_utc: pd.DataFrame, ticker: str, cfg) -> pd.DataFrame:
    """Uma linha a cada N minutos, de janela_inicio a janela_fim (inclusive), por pregão com dado.

    A barra que falta repete o último valor do mesmo dia. As barras do próprio dia entre a abertura do
    pregão contínuo e o início da janela servem de ponto de partida; as do leilão de abertura, não.
    Nada é filtrado.
    """
    if df_utc.empty:
        return pd.DataFrame({"data_hora": pd.Series(dtype="datetime64[ns]"), "ticker": pd.Series(dtype=str),
                             "compra": pd.Series(dtype=float), "venda": pd.Series(dtype=float),
                             "ultimo": pd.Series(dtype=float)})
    loc = df_utc.tz_convert(TZ).tz_localize(None)
    passo = pd.Timedelta(minutes=cfg["intervalo_min"])
    partes = []
    for dia, g in loc.groupby(loc.index.date):
        ta = pd.Timestamp(dt.datetime.combine(dia, cfg["abertura_continuo"]))
        t0 = pd.Timestamp(dt.datetime.combine(dia, cfg["janela_inicio"]))
        t1 = pd.Timestamp(dt.datetime.combine(dia, cfg["janela_fim"]))
        g = g[(g.index >= ta) & (g.index <= t1)]
        if g.empty:
            continue  # nenhuma barra no pregão: dia sem negociação (feriado, por exemplo)
        grade = pd.date_range(t0, t1, freq=passo)
        r = g.reindex(g.index.union(grade)).ffill().reindex(grade)  # ffill só dentro do dia
        r = r.rename(columns={CAMPO_BID: "compra", CAMPO_ASK: "venda", CAMPO_ULTIMO: "ultimo"})
        r.insert(0, "ticker", ticker)
        r.index.name = "data_hora"
        partes.append(r.reset_index())
    return pd.concat(partes, ignore_index=True) if partes else montar_grade(_vazio(), ticker, cfg)


# --------------------------------------------------------------------------------------------
def versao_biblioteca(ld) -> str:
    for nome in ("lseg-data", "refinitiv-data"):
        try:
            from importlib.metadata import version
            return f"{nome} {version(nome)}"
        except Exception:  # noqa: BLE001
            continue
    return getattr(ld, "__name__", "?")


def main():
    args = sys.argv[1:]
    if args and args[0].lower().endswith(".csv"):  # python spread_etfs.py saida\<bruto>.csv [planilha.xlsx]
        so_spread(Path(args[0]), Path(args[1]) if len(args) > 1 else PASTA / "inputs_spread_etfs.xlsx")
        return
    entrada = Path(args[0]) if args else PASTA / "inputs_spread_etfs.xlsx"
    etfs, cfg = ler_entrada(entrada)
    log(f"{len(etfs)} ETFs: {', '.join(etfs['ticker'])} | {cfg['data_inicio']:%d/%m/%Y} a {cfg['data_fim']:%d/%m/%Y} "
        f"| barras de {cfg['intervalo_min']} min | {cfg['janela_inicio']:%H:%M} a {cfg['janela_fim']:%H:%M}")

    ld = importar_api()
    abrir_sessao(ld)
    tabelas, resumo = [], []
    try:
        for etf in etfs.itertuples():
            log(f"{etf.ticker} ({etf.ric})")
            bruto = baixar_periodo(ld, etf.ric, cfg)
            primeira = bruto.index.min() if len(bruto) else None
            inicio_hist, como = sondar_inicio(ld, etf.ric, cfg, primeira)
            t = montar_grade(bruto, etf.ticker, cfg)
            tabelas.append(t)
            sem_valor = int(t[["compra", "venda"]].isna().any(axis=1).sum()) if len(t) else 0
            resumo.append({
                "ticker": etf.ticker, "ric": etf.ric,
                "primeira_barra_periodo": primeira.tz_convert(TZ).strftime("%d/%m/%Y %H:%M") if primeira is not None else "-",
                "historico_desde": inicio_hist.tz_convert(TZ).strftime("%d/%m/%Y") if inicio_hist is not None else "-",
                "como": como,
                "pregoes": t["data_hora"].dt.date.nunique() if len(t) else 0,
                "linhas": len(t),
                "linhas_sem_compra_ou_venda": sem_valor,
                "ultimo_vazio": bool(len(t)) and bool(t["ultimo"].isna().all()),
            })
            log(f"   {etf.ticker}: {resumo[-1]['pregoes']} pregões, {len(t)} linhas; histórico desde {resumo[-1]['historico_desde']}")
    finally:
        try:
            ld.close_session()
        except Exception:  # noqa: BLE001
            pass

    saida = PASTA / cfg["pasta_saida"]
    saida.mkdir(exist_ok=True)
    nome = cfg["arquivo_saida"] or f"etfs_bid_ask_{cfg['intervalo_min']}min_{cfg['data_inicio']:%Y%m%d}_{cfg['data_fim']:%Y%m%d}"
    nome = nome[:-4] if nome.lower().endswith(".csv") else nome
    base = pd.concat(tabelas, ignore_index=True)
    if not cfg["incluir_ultimo"]:
        base = base.drop(columns="ultimo")
    bruto_csv = base.copy()
    bruto_csv["data_hora"] = bruto_csv["data_hora"].dt.strftime("%Y-%m-%d %H:%M:%S")
    arq_csv = saida / f"{nome}.csv"
    bruto_csv.to_csv(arq_csv, index=False, sep=",", decimal=".", encoding="utf-8", float_format="%.2f")
    log(f"Pronto: {arq_csv}")
    arq_spread = None
    if cfg["gerar_spread"] and len(base):
        arq_spread = gravar_spread(base, list(etfs["ticker"]), cfg, saida / f"spread_{nome}.xlsx")
        log(f"Spread calculado: {arq_spread}")
    arq_txt = saida / f"relatorio_{nome}.txt"
    log(f"Relatório: {arq_txt}")
    arq_txt.write_text(relatorio(cfg, resumo, ld, arq_csv, arq_spread), encoding="utf-8")


def so_spread(arq_csv: Path, entrada: Path):
    """Refaz só o Excel do spread a partir de um CSV bruto já gerado (sem chamar a API)."""
    _, cfg = ler_entrada(entrada)
    base = pd.read_csv(arq_csv)
    ordem = list(dict.fromkeys(base["ticker"]))
    arq = gravar_spread(base, ordem, cfg, arq_csv.with_name(f"spread_{arq_csv.stem}.xlsx"))
    log(f"Spread calculado: {arq}")


# --------------------------------------------------------------------------------------------
# Spread calculado (arquivo à parte; o CSV bruto não muda)
# --------------------------------------------------------------------------------------------
def calcular_spread(base: pd.DataFrame, min_validos: int):
    """A partir do CSV bruto (data_hora, ticker, compra, venda): spread de cada barra e agregados.

    Barra válida: compra > 0, venda > 0 e venda > compra. As outras ficam fora da conta (no bruto, não).
    """
    b = base.copy()
    b["data_hora"] = pd.to_datetime(b["data_hora"])
    b["data"] = b["data_hora"].dt.normalize()
    b["horario"] = b["data_hora"].dt.strftime("%H:%M")
    valido = (b["compra"] > 0) & (b["venda"] > 0) & (b["venda"] > b["compra"])
    meio = (b["compra"] + b["venda"]) / 2
    b["spread_bps"] = ((b["venda"] - b["compra"]) / meio * 1e4).where(valido)
    b["spread_rs"] = (b["venda"] - b["compra"]).where(valido)
    b["valido"] = valido

    g = b.groupby(["ticker", "data"])
    d = g.agg(barras=("valido", "size"), barras_validas=("valido", "sum"),
              spread_medio_bps=("spread_bps", "mean"), spread_mediano_bps=("spread_bps", "median"),
              spread_min_bps=("spread_bps", "min"), spread_max_bps=("spread_bps", "max"),
              spread_medio_rs=("spread_rs", "mean")).reset_index()
    d["spread_p90_bps"] = g["spread_bps"].quantile(0.9).values
    d["pct_barras_validas"] = d["barras_validas"] / d["barras"]
    pouco = d["barras_validas"] < min_validos
    for c in ("spread_medio_bps", "spread_mediano_bps", "spread_min_bps", "spread_max_bps", "spread_p90_bps", "spread_medio_rs"):
        d.loc[pouco, c] = np.nan
    d = d[["ticker", "data", "barras", "barras_validas", "pct_barras_validas", "spread_medio_bps", "spread_mediano_bps",
           "spread_p90_bps", "spread_min_bps", "spread_max_bps", "spread_medio_rs"]]
    return b, d


def gravar_spread(base: pd.DataFrame, ordem: list[str], cfg, arq: Path) -> Path:
    b, d = calcular_spread(base, cfg["min_barras_validas"])
    ordem = [t for t in ordem if t in set(d["ticker"])]
    ok = d[d["spread_medio_bps"].notna()]

    def ult(s, n=21):
        return s.sort_index().tail(n).mean()

    linhas = []
    for t in ordem:
        x = ok[ok["ticker"] == t].set_index("data")
        todos = d[d["ticker"] == t]
        linhas.append({
            "ETF": t,
            "Pregões": len(x),
            "Spread médio (bps)": x["spread_medio_bps"].mean(),
            "Spread médio, últimos 21 pregões (bps)": ult(x["spread_medio_bps"]),
            "Mediana dos dias (bps)": x["spread_medio_bps"].median(),
            "P10 dos dias (bps)": x["spread_medio_bps"].quantile(0.1),
            "P90 dos dias (bps)": x["spread_medio_bps"].quantile(0.9),
            "Spread médio (R$)": x["spread_medio_rs"].mean(),
            "% das barras com cotação válida": todos["barras_validas"].sum() / todos["barras"].sum() if len(todos) else np.nan,
            "Primeiro pregão": x.index.min(),
            "Último pregão": x.index.max(),
        })
    resumo = pd.DataFrame(linhas)

    diario = ok.pivot_table(index="data", columns="ticker", values="spread_medio_bps").reindex(columns=ordem)
    diario.index.name = "Data"
    mensal = diario.groupby(diario.index.to_period("M")).mean()
    mensal.index = mensal.index.to_timestamp()
    mensal.index.name = "Mês"
    bv = b[b["valido"]]
    perfil = bv.pivot_table(index="horario", columns="ticker", values="spread_bps", aggfunc="mean").reindex(columns=ordem)
    perfil.index.name = "Horário"

    det = d.rename(columns={
        "ticker": "ETF", "data": "Data", "barras": "Barras", "barras_validas": "Barras válidas",
        "pct_barras_validas": "% válidas", "spread_medio_bps": "Spread médio (bps)",
        "spread_mediano_bps": "Mediana (bps)", "spread_p90_bps": "P90 (bps)", "spread_min_bps": "Mínimo (bps)",
        "spread_max_bps": "Máximo (bps)", "spread_medio_rs": "Spread médio (R$)"})
    det["ETF"] = pd.Categorical(det["ETF"], ordem, ordered=True)
    det = det.sort_values(["ETF", "Data"])

    notas = pd.DataFrame({"Notas": [
        "Spread de cada barra = (venda - compra) / ponto médio x 10.000, em bps, com a cotação do fim da barra "
        f"de {cfg['intervalo_min']} minutos, de {cfg['janela_inicio']:%H:%M} a {cfg['janela_fim']:%H:%M}.",
        "Barra válida: compra > 0, venda > 0 e venda > compra. Livro cruzado/travado, oferta zerada ou vazia ficam "
        "fora da conta (no CSV bruto continuam).",
        f"Spread do dia = média das barras válidas do pregão; dia com menos de {cfg['min_barras_validas']} barras "
        "válidas fica sem spread.",
        "Resumo e Mensal: média simples dos spreads diários (cada pregão pesa igual).",
        "Perfil intradiário: média do spread de cada horário, em todos os pregões do período.",
        "A amostra a cada 5 minutos aproxima a média ponderada pelo tempo; não pega spreads que abrem e fecham "
        "dentro da barra.",
        "Fonte: LSEG Workspace (BID/ASK). Dado bruto da LSEG: não redistribuir.",
        f"Gerado em {dt.datetime.now():%d/%m/%Y %H:%M}.",
    ]})

    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        resumo.to_excel(w, sheet_name="Resumo", index=False)
        mensal.reset_index().to_excel(w, sheet_name="Mensal (bps)", index=False)
        diario.reset_index().to_excel(w, sheet_name="Diário (bps)", index=False)
        perfil.reset_index().to_excel(w, sheet_name="Perfil intradiário (bps)", index=False)
        det.to_excel(w, sheet_name="Diário detalhado", index=False)
        notas.to_excel(w, sheet_name="Notas", index=False)
        _formatar_spread(w)
    return arq


def _formatar_spread(w) -> None:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import ColorScaleRule
    navy, fino = PatternFill("solid", fgColor="0B1F3A"), Side(style="thin", color="BFBFBF")
    for ws in w.book.worksheets:
        ws.sheet_view.showGridLines = False
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = navy
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 32
        if ws.title == "Notas":
            ws.column_dimensions["A"].width = 130
            continue
        ws.freeze_panes = "B2"
        for col in ws.iter_cols(min_row=1):
            nome = str(col[0].value or "")
            letra = col[0].column_letter
            ws.column_dimensions[letra].width = 13 if col[0].column > 1 else 14
            if ws.title == "Resumo":
                ws.column_dimensions[letra].width = 16
            for c in col[1:]:
                c.border = Border(bottom=fino)
                if col[0].column == 1 and ws.title in ("Mensal (bps)", "Diário (bps)", "Diário detalhado"):
                    c.number_format = "mmm/yy" if ws.title.startswith("Mensal") else "dd/mm/yyyy"
                elif isinstance(c.value, (int, float)) or c.value is None:
                    if "%" in nome:
                        c.number_format = "0.0%"
                    elif "R$" in nome:
                        c.number_format = "0.000"
                    elif "Pregões" in nome or "Barras" in nome:
                        c.number_format = "0"
                    else:
                        c.number_format = "0.00"
                elif "pregão" in nome.lower():
                    c.number_format = "dd/mm/yyyy"
                if ws.title == "Diário detalhado" and nome == "Data":
                    c.number_format = "dd/mm/yyyy"
            if ws.title == "Resumo" and "pregão" in nome.lower():
                for c in col[1:]:
                    c.number_format = "dd/mm/yyyy"
        # escala de cor (verde = spread menor) nas tabelas data/horário × ETF
        if ws.title in ("Mensal (bps)", "Diário (bps)", "Perfil intradiário (bps)") and ws.max_column > 1:
            from openpyxl.utils import get_column_letter
            for j in range(2, ws.max_column + 1):
                L = get_column_letter(j)
                ws.conditional_formatting.add(f"{L}2:{L}{ws.max_row}", ColorScaleRule(
                    start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50,
                    mid_color="FFEB84", end_type="max", end_color="F8696B"))


def relatorio(cfg, resumo, ld, arq_csv, arq_spread=None) -> str:
    r = pd.DataFrame(resumo)
    campos = campos_da_chamada(cfg)
    linhas = [
        "Barras intraday de ETFs (LSEG Workspace)",
        f"Gerado em {dt.datetime.now():%d/%m/%Y %H:%M}. Arquivo bruto: {arq_csv.name}"
        + (f"; spread calculado: {arq_spread.name}" if arq_spread else ""),
        "",
        "1. Desde quando o histórico intradiário existe",
    ]
    for x in resumo:
        linhas.append(f"   {x['ticker']}: {x['historico_desde']} ({x['como']}); primeira barra no período pedido: "
                      f"{x['primeira_barra_periodo']}")
    linhas += [
        f"   Período pedido: {cfg['data_inicio']:%d/%m/%Y} a {cfg['data_fim']:%d/%m/%Y}.",
        "",
        "2. Chamada usada",
        f"   Biblioteca: {versao_biblioteca(ld)}, função get_history.",
        f"   universe = <RIC> (ex.: {resumo[0]['ric'] if resumo else 'BOVA11.SA'})",
        f"   fields = {campos}" + ("   (TRDPRC_1 = último preço)" if cfg["incluir_ultimo"] else ""),
        f"   interval = \"{cfg['intervalo_api']}\"",
        "   start/end = um mês calendário por chamada, em UTC (o mês é dividido se a resposta vier cortada).",
        "   Valor de cada barra: o último do intervalo, como a API devolve.",
        "   Horário: a API devolve em UTC; o CSV está no horário de Brasília.",
        "",
        "3. Tratamento",
        f"   Uma linha a cada {cfg['intervalo_min']} minutos, de {cfg['janela_inicio']:%H:%M} a {cfg['janela_fim']:%H:%M} "
        "(inclusive), só nos pregões com dado.",
        "   Barra que falta: repete o último valor do mesmo dia, nunca do dia anterior. Barras do próprio dia entre",
        f"   {cfg['abertura_continuo']:%H:%M} e {cfg['janela_inicio']:%H:%M} servem de ponto de partida; as do leilão de "
        "abertura, não.",
        "   Se o dia ainda não tinha nenhuma cotação até aquele horário, a linha fica com compra/venda vazios.",
        "   Carimbo de hora: o da barra como a API devolve. Confirmar na documentação da LSEG se ele marca o",
        "   início ou o fim do intervalo (muda em 5 minutos a leitura de cada linha).",
        "   Nenhum outro filtro (livro cruzado, oferta zerada etc. vão como vieram).",
        "",
        "4. Linhas por ETF",
    ]
    for x in resumo:
        linhas.append(f"   {x['ticker']}: {x['pregoes']} pregões, {x['linhas']} linhas, "
                      f"{x['linhas_sem_compra_ou_venda']} linhas sem compra ou venda"
                      + ("; último preço veio vazio" if cfg["incluir_ultimo"] and x["ultimo_vazio"] else ""))
    linhas += [
        "",
        "5. Campo diário pronto de spread médio",
        "   Não verificado pelo script. Buscar 'spread' no Data Item Browser (DIB) do Workspace e conferir na",
        "   definição se é média do pregão ou foto do fechamento.",
        "",
        "Log da rodada:",
    ] + [f"   {l}" for l in LOG]
    if len(r) and (r["pregoes"] == 0).any():
        linhas.insert(4, "   ATENÇÃO: ETF sem nenhum dado no período. Confira o RIC.")
    return "\n".join(linhas) + "\n"


if __name__ == "__main__":
    main()
