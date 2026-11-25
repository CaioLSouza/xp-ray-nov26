"""
Barras intraday de BID, ASK, quantidade no topo do livro (BIDSIZE, ASKSIZE) e último preço (TRDPRC_1) de ETFs
da B3, pela LSEG (Refinitiv Workspace), no formato pedido pelo Research (versão de 2026-11-25, 2ª rodada).
Gera também um Excel à parte com o spread calculado.

Uso (no PC com o Workspace aberto e logado):
    python spread_etfs.py                       # lê inputs_spread_etfs.xlsx (período da aba Parametros)
    python spread_etfs.py --mes-anterior        # rotina mensal: só o mês calendário anterior
    python spread_etfs.py outra_planilha.xlsx   # outra planilha de entrada (combina com --mes-anterior)
    python spread_etfs.py saida\\<bruto>.csv     # refaz só o Excel do spread, sem chamar a API

O que faz:
  1. Lê os ETFs e os parâmetros na planilha de entrada.
  2. Para cada ETF, baixa as barras (padrão: 5 minutos) dos campos da aba Parametros, um mês por chamada.
     Usa o último valor de cada barra, como a API devolve. Se a resposta vier cortada pelo limite de
     linhas, divide o mês sozinho. Os meses já fechados ficam em cache.
  3. Converte o horário de UTC para Brasília e monta, para cada pregão, uma linha a cada 5 minutos de
     10:15 a 16:45 (inclusive), sem buracos. Barra sem atualização repete o último valor do MESMO dia (nunca
     do dia anterior); se o dia ainda não tinha oferta, a linha fica vazia (nunca zero). Não filtra mais nada.
  4. Grava saida\\etfs_bid_ask_5min_AAAAMMDD_AAAAMMDD.csv
     (data_hora, ticker, compra, venda, qtd_compra, qtd_venda, ultimo) e um relatorio_*.txt com a data de
     execução, os campos usados na chamada, desde quando há histórico e as falhas (ETF sem dado, dia faltando,
     campo vazio).

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
# coluna do CSV -> campo da API (padrão; a aba Parametros pode trocar o nome do campo)
CAMPOS_PADRAO = {"compra": "BID", "venda": "ASK", "qtd_compra": "BIDSIZE", "qtd_venda": "ASKSIZE",
                 "ultimo": "TRDPRC_1"}
COLUNAS = list(CAMPOS_PADRAO)          # ordem das colunas de valor no CSV
COLS_QTD = ("qtd_compra", "qtd_venda")  # saem como inteiro (cotas)
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


def mes_anterior(hoje: pd.Timestamp | None = None):
    hoje = (hoje or pd.Timestamp.today()).normalize()
    fim = hoje.replace(day=1) - pd.Timedelta(days=1)
    return fim.replace(day=1), fim


def ler_entrada(caminho: Path, forcar_mes_anterior: bool = False):
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
    periodo = str(p("periodo", "datas")).strip().lower().replace(" ", "_").replace("ê", "e")
    mensal = forcar_mes_anterior or periodo in ("mes_anterior", "mensal")
    cfg = {
        "data_inicio": pd.Timestamp(p("data_inicio", ontem - pd.Timedelta(days=365))).normalize(),
        "data_fim": pd.Timestamp(p("data_fim", ontem)).normalize(),
        "mensal": mensal,
        "intervalo_min": int(p("intervalo_min", 5)),
        "janela_inicio": _hora(p("janela_inicio", "10:15")),
        "janela_fim": _hora(p("janela_fim", "16:45")),
        "abertura_continuo": _hora(p("abertura_continuo", "10:00")),
        "campos": {col: str(p(f"campo_{col}", api)).strip().upper() for col, api in CAMPOS_PADRAO.items()},
        "gerar_spread": _sim(p("gerar_spread", "S")),
        "min_barras_validas": int(p("min_barras_validas", 20)),
        "sondar_historico": _sim(p("sondar_historico", "S")) and not mensal,
        "pausa_seg": float(p("pausa_seg", 0.5)),
        "usar_cache": _sim(p("usar_cache", "S")),
        "pasta_saida": str(p("pasta_saida", "saida")),
        "arquivo_saida": str(p("arquivo_saida", "")).strip(),
    }
    if mensal:
        cfg["data_inicio"], cfg["data_fim"] = mes_anterior()
        cfg["arquivo_saida"] = ""  # o nome sai do mês, para uma rodada não sobrescrever a outra
    if cfg["data_fim"] < cfg["data_inicio"]:
        raise SystemExit("data_fim antes de data_inicio na aba Parametros.")
    if not cfg["abertura_continuo"] <= cfg["janela_inicio"] < cfg["janela_fim"]:
        raise SystemExit("Na aba Parametros, a ordem tem de ser abertura_continuo <= janela_inicio < janela_fim.")
    if len(set(cfg["campos"].values())) != len(cfg["campos"]):
        raise SystemExit(f"Há campo repetido na aba Parametros: {cfg['campos']}")
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


def _vazio(campos) -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype=float) for c in campos}, index=pd.DatetimeIndex([], tz="UTC"))


def _normalizar(df: pd.DataFrame | None, campos) -> pd.DataFrame:
    """Põe a resposta da API no formato fixo: índice em UTC e uma coluna por campo (vazia se não veio).

    Os dois primeiros campos (compra e venda) são obrigatórios: se não vierem, para com erro.
    """
    if df is None or len(df) == 0:
        return _vazio(campos)
    df = df.copy()
    if isinstance(df.columns, pd.MultiIndex):  # alguns retornos vêm com (RIC, campo)
        df.columns = [c[-1] for c in df.columns]
    df.columns = [str(c).upper() for c in df.columns]
    faltando = [c for c in campos[:2] if c not in df.columns]
    if faltando:
        raise RuntimeError(f"A resposta não trouxe {faltando}. Colunas recebidas: {list(df.columns)}")
    for c in campos:
        if c not in df.columns:
            df[c] = np.nan
    df = df[list(campos)].apply(pd.to_numeric, errors="coerce")
    idx = pd.to_datetime(df.index)
    idx = idx.tz_localize("UTC") if idx.tz is None else idx.tz_convert("UTC")  # a LSEG devolve em UTC
    df.index = idx
    return df[~df.index.isna()].sort_index()


def chamar_api(ld, ric, campos, intervalo, ini_utc, fim_utc, pausa) -> pd.DataFrame:
    for tentativa in range(1, 4):
        try:
            df = ld.get_history(
                universe=ric, fields=list(campos), interval=intervalo,
                start=ini_utc.strftime("%Y-%m-%dT%H:%M:%S"), end=fim_utc.strftime("%Y-%m-%dT%H:%M:%S"),
            )
            time.sleep(pausa)
            return _normalizar(df, campos)
        except RuntimeError:
            raise
        except Exception as e:  # noqa: BLE001
            texto = str(e).lower()
            if "no data" in texto or "not found" in texto or "empty" in texto:
                time.sleep(pausa)
                return _vazio(campos)
            espera = 5 * tentativa
            log(f"   erro na API ({ric}, {ini_utc:%Y-%m-%d}): {e}. Nova tentativa em {espera}s.")
            time.sleep(espera)
    raise RuntimeError(f"a API falhou 3 vezes ({ric}, {ini_utc:%d/%m/%Y} a {fim_utc:%d/%m/%Y})")


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


def campos_da_chamada(cfg) -> list[str]:
    return [cfg["campos"][c] for c in COLUNAS]


def baixar_periodo(ld, ric, cfg) -> pd.DataFrame:
    cache_dir = PASTA / "cache"
    cache_dir.mkdir(exist_ok=True)
    hoje = pd.Timestamp.today().normalize()
    campos = campos_da_chamada(cfg)
    partes = []
    for a, b in meses(cfg["data_inicio"], cfg["data_fim"]):
        arq = cache_dir / f"{ric.replace('.', '_')}_{cfg['intervalo_api']}_{'+'.join(campos)}_{a:%Y%m%d}_{b:%Y%m%d}.csv"
        if cfg["usar_cache"] and arq.exists() and b <= hoje:
            # passa pelo mesmo formato da API: um mês vazio no cache volta como texto e estragaria a concatenação
            df = _normalizar(pd.read_csv(arq, index_col=0), campos)
        else:
            df = baixar(ld, ric, campos, cfg["intervalo_api"],
                        a.tz_localize(TZ).tz_convert("UTC"), b.tz_localize(TZ).tz_convert("UTC"), cfg["pausa_seg"])
            if b <= hoje:  # só guarda mês fechado
                df.to_csv(arq)
        log(f"   {ric} {a:%d/%m/%Y} a {b - pd.Timedelta(days=1):%d/%m/%Y}: {len(df)} barras")
        partes.append(df)
    out = pd.concat(partes) if partes else _vazio(campos)
    return out[~out.index.duplicated(keep="last")].sort_index()


def sondar_inicio(ld, ric, cfg, primeira_no_periodo):
    """Se o período pedido já começa com dado, volta mês a mês para achar desde quando existe histórico."""
    if primeira_no_periodo is None:
        return None, "sem dado no período pedido"
    if primeira_no_periodo.tz_convert(TZ).date() > (cfg["data_inicio"] + pd.Timedelta(days=7)).date():
        return primeira_no_periodo, "o histórico começa dentro do período pedido"
    if not cfg["sondar_historico"]:
        return None, ("rotina mensal, sem sondagem" if cfg["mensal"]
                      else "existe pelo menos desde o início do período (sondagem desligada)")
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
def _grade_vazia() -> pd.DataFrame:
    d = {"data_hora": pd.Series(dtype="datetime64[ns]"), "ticker": pd.Series(dtype=str)}
    d.update({c: pd.Series(dtype=float) for c in COLUNAS})
    return pd.DataFrame(d)


def montar_grade(df_utc: pd.DataFrame, ticker: str, cfg) -> pd.DataFrame:
    """Uma linha a cada N minutos, de janela_inicio a janela_fim (inclusive), por pregão com dado.

    Barra sem atualização repete o último valor do mesmo dia, campo a campo. As barras do próprio dia entre a
    abertura do pregão contínuo e o início da janela servem de ponto de partida; as do leilão de abertura, não.
    Antes da primeira oferta do dia a linha fica vazia (nunca zero). Nada é filtrado.
    """
    if df_utc.empty:
        return _grade_vazia()
    nomes = {api: col for col, api in cfg["campos"].items()}
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
        r = r.rename(columns=nomes)[COLUNAS]
        r.insert(0, "ticker", ticker)
        r.index.name = "data_hora"
        partes.append(r.reset_index())
    return pd.concat(partes, ignore_index=True) if partes else _grade_vazia()


def dias_faltando(tabelas: dict[str, pd.DataFrame], cfg):
    """Dias faltando por ETF e dias úteis sem dado em nenhum ETF.

    O calendário de pregões é o conjunto dos dias em que pelo menos um ETF teve dado. Um ETF "perde" o dia quando
    o dia está nesse calendário, cai depois do primeiro pregão dele no período e ele não teve barra na janela.
    """
    dias = {t: set(x["data_hora"].dt.date) for t, x in tabelas.items()}
    calendario = set().union(*dias.values()) if dias else set()
    # antes do 1º pregão com dado é "antes do histórico" (seção 2 do relatório), não falha nem feriado
    uteis = {d.date() for d in pd.bdate_range(min(calendario), cfg["data_fim"])} if calendario else set()
    sem_nenhum = sorted(uteis - calendario)
    por_etf = {}
    for t, ds in dias.items():
        if ds:
            falta = sorted(d for d in calendario if d >= min(ds) and d not in ds)
            if falta:
                por_etf[t] = falta
    return por_etf, sem_nenhum


def _datas(lista, n=12) -> str:
    txt = ", ".join(f"{d:%d/%m/%Y}" for d in lista[:n])
    return txt + (f" e mais {len(lista) - n}" if len(lista) > n else "")


# --------------------------------------------------------------------------------------------
def versao_biblioteca(ld) -> str:
    for nome in ("lseg-data", "refinitiv-data"):
        try:
            from importlib.metadata import version
            return f"{nome} {version(nome)}"
        except Exception:  # noqa: BLE001
            continue
    return getattr(ld, "__name__", "?")


def gravar_csv(base: pd.DataFrame, arq: Path) -> None:
    """CSV UTF-8, vírgula, ponto decimal, sem índice; preço com 2 casas, quantidade inteira, vazio = vazio."""
    x = base[["data_hora", "ticker"] + COLUNAS].copy()
    x["data_hora"] = x["data_hora"].dt.strftime("%Y-%m-%d %H:%M:%S")
    for c in COLUNAS:
        x[c] = pd.to_numeric(x[c], errors="coerce").astype(float)
    for c in COLS_QTD:
        x[c] = x[c].round().astype("Int64")
    x.to_csv(arq, index=False, sep=",", decimal=".", encoding="utf-8", float_format="%.2f")


def main():
    args = [a for a in sys.argv[1:]]
    forcar_mensal = any(a.lower() in ("--mes-anterior", "--mensal") for a in args)
    args = [a for a in args if not a.startswith("--")]
    if args and args[0].lower().endswith(".csv"):  # python spread_etfs.py saida\<bruto>.csv [planilha.xlsx]
        so_spread(Path(args[0]), Path(args[1]) if len(args) > 1 else PASTA / "inputs_spread_etfs.xlsx")
        return
    entrada = Path(args[0]) if args else PASTA / "inputs_spread_etfs.xlsx"
    etfs, cfg = ler_entrada(entrada, forcar_mensal)
    executado = dt.datetime.now()
    log(f"{len(etfs)} ETFs: {', '.join(etfs['ticker'])}")
    log(f"Período {cfg['data_inicio']:%d/%m/%Y} a {cfg['data_fim']:%d/%m/%Y}"
        f"{' (rotina mensal: mês anterior)' if cfg['mensal'] else ''} | barras de {cfg['intervalo_min']} min | "
        f"{cfg['janela_inicio']:%H:%M} a {cfg['janela_fim']:%H:%M} | campos {campos_da_chamada(cfg)}")

    ld = importar_api()
    abrir_sessao(ld)
    tabelas, resumo = {}, []
    try:
        for etf in etfs.itertuples():
            log(f"{etf.ticker} ({etf.ric})")
            item = {"ticker": etf.ticker, "ric": etf.ric, "erro": "", "primeira_barra_periodo": "-",
                    "historico_desde": "-", "como": "", "pregoes": 0, "linhas": 0,
                    "linhas_sem_compra_ou_venda": 0, "campos_vazios": []}
            try:
                bruto = baixar_periodo(ld, etf.ric, cfg)
                primeira = bruto.index.min() if len(bruto) else None
                inicio_hist, como = sondar_inicio(ld, etf.ric, cfg, primeira)
                t = montar_grade(bruto, etf.ticker, cfg)
            except Exception as e:  # noqa: BLE001  (um ETF com erro não derruba os outros)
                item["erro"] = str(e)
                log(f"   ERRO em {etf.ticker}: {e}")
                resumo.append(item)
                continue
            tabelas[etf.ticker] = t
            item.update({
                "primeira_barra_periodo": primeira.tz_convert(TZ).strftime("%d/%m/%Y %H:%M") if primeira is not None else "-",
                "historico_desde": inicio_hist.tz_convert(TZ).strftime("%d/%m/%Y") if inicio_hist is not None else "-",
                "como": como,
                "pregoes": t["data_hora"].dt.date.nunique() if len(t) else 0,
                "linhas": len(t),
                "linhas_sem_compra_ou_venda": int(t[["compra", "venda"]].isna().any(axis=1).sum()) if len(t) else 0,
                "campos_vazios": [c for c in COLUNAS[2:] if len(t) and t[c].isna().all()],
            })
            resumo.append(item)
            log(f"   {etf.ticker}: {item['pregoes']} pregões, {len(t)} linhas; histórico desde {item['historico_desde']}")
    finally:
        try:
            ld.close_session()
        except Exception:  # noqa: BLE001
            pass

    saida = PASTA / cfg["pasta_saida"]
    saida.mkdir(exist_ok=True)
    nome = cfg["arquivo_saida"] or f"etfs_bid_ask_{cfg['intervalo_min']}min_{cfg['data_inicio']:%Y%m%d}_{cfg['data_fim']:%Y%m%d}"
    nome = nome[:-4] if nome.lower().endswith(".csv") else nome
    base = pd.concat(list(tabelas.values()), ignore_index=True) if tabelas else _grade_vazia()
    arq_csv = saida / f"{nome}.csv"
    gravar_csv(base, arq_csv)
    log(f"Pronto: {arq_csv}")
    arq_spread = None
    if cfg["gerar_spread"] and len(base):
        arq_spread = gravar_spread(base, list(etfs["ticker"]), cfg, saida / f"spread_{nome}.xlsx")
        log(f"Spread calculado: {arq_spread}")
    faltas = dias_faltando(tabelas, cfg)
    arq_txt = saida / f"relatorio_{nome}.txt"
    log(f"Relatório: {arq_txt}")
    texto, n_falhas = relatorio(cfg, resumo, faltas, ld, arq_csv, arq_spread, executado)
    arq_txt.write_text(texto, encoding="utf-8")
    log(f"{n_falhas} falha(s) registrada(s) no relatório." if n_falhas else "Sem falhas.")


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
    """A partir do CSV bruto (data_hora, ticker, compra, venda[, qtd_compra, qtd_venda]): spread de cada barra
    e agregados por pregão.

    Barra válida: compra > 0, venda > 0 e venda > compra. As outras ficam fora da conta (no bruto, não).
    Profundidade no topo = média das quantidades na melhor compra e na melhor venda, nas barras válidas.
    """
    b = base.copy()
    b["data_hora"] = pd.to_datetime(b["data_hora"])
    b["data"] = b["data_hora"].dt.normalize()
    b["horario"] = b["data_hora"].dt.strftime("%H:%M")
    valido = (b["compra"] > 0) & (b["venda"] > 0) & (b["venda"] > b["compra"])
    meio = (b["compra"] + b["venda"]) / 2
    b["spread_bps"] = ((b["venda"] - b["compra"]) / meio * 1e4).where(valido)
    b["spread_rs"] = (b["venda"] - b["compra"]).where(valido)
    tem_qtd = all(c in b.columns for c in COLS_QTD) and b[list(COLS_QTD)].notna().any().all()
    if tem_qtd:
        b["prof_cotas"] = ((b["qtd_compra"] + b["qtd_venda"]) / 2).where(valido)
        b["prof_rs_mil"] = ((b["qtd_compra"] * b["compra"] + b["qtd_venda"] * b["venda"]) / 2 / 1e3).where(valido)
    else:
        b["prof_cotas"] = np.nan
        b["prof_rs_mil"] = np.nan
    b["valido"] = valido

    g = b.groupby(["ticker", "data"])
    d = g.agg(barras=("valido", "size"), barras_validas=("valido", "sum"),
              spread_medio_bps=("spread_bps", "mean"), spread_mediano_bps=("spread_bps", "median"),
              spread_min_bps=("spread_bps", "min"), spread_max_bps=("spread_bps", "max"),
              spread_medio_rs=("spread_rs", "mean"), prof_cotas=("prof_cotas", "mean"),
              prof_rs_mil=("prof_rs_mil", "mean")).reset_index()
    d["spread_p90_bps"] = g["spread_bps"].quantile(0.9).values
    d["pct_barras_validas"] = d["barras_validas"] / d["barras"]
    pouco = d["barras_validas"] < min_validos
    for c in ("spread_medio_bps", "spread_mediano_bps", "spread_min_bps", "spread_max_bps", "spread_p90_bps",
              "spread_medio_rs", "prof_cotas", "prof_rs_mil"):
        d.loc[pouco, c] = np.nan
    d = d[["ticker", "data", "barras", "barras_validas", "pct_barras_validas", "spread_medio_bps", "spread_mediano_bps",
           "spread_p90_bps", "spread_min_bps", "spread_max_bps", "spread_medio_rs", "prof_cotas", "prof_rs_mil"]]
    return b, d, tem_qtd


def gravar_spread(base: pd.DataFrame, ordem: list[str], cfg, arq: Path) -> Path:
    b, d, tem_qtd = calcular_spread(base, cfg["min_barras_validas"])
    ordem = [t for t in ordem if t in set(d["ticker"])]
    ok = d[d["spread_medio_bps"].notna()]

    def ult(s, n=21):
        return s.sort_index().tail(n).mean()

    linhas = []
    for t in ordem:
        x = ok[ok["ticker"] == t].set_index("data")
        todos = d[d["ticker"] == t]
        linha = {
            "ETF": t,
            "Pregões": len(x),
            "Spread médio (bps)": x["spread_medio_bps"].mean(),
            "Spread médio, últimos 21 pregões (bps)": ult(x["spread_medio_bps"]),
            "Mediana dos dias (bps)": x["spread_medio_bps"].median(),
            "P10 dos dias (bps)": x["spread_medio_bps"].quantile(0.1),
            "P90 dos dias (bps)": x["spread_medio_bps"].quantile(0.9),
            "Spread médio (R$)": x["spread_medio_rs"].mean(),
        }
        if tem_qtd:
            linha["Profundidade média no topo (cotas)"] = x["prof_cotas"].mean()
            linha["Profundidade média no topo (R$ mil)"] = x["prof_rs_mil"].mean()
        linha.update({
            "% das barras com cotação válida": todos["barras_validas"].sum() / todos["barras"].sum() if len(todos) else np.nan,
            "Primeiro pregão": x.index.min(),
            "Último pregão": x.index.max(),
        })
        linhas.append(linha)
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
        "spread_max_bps": "Máximo (bps)", "spread_medio_rs": "Spread médio (R$)",
        "prof_cotas": "Profundidade no topo (cotas)", "prof_rs_mil": "Profundidade no topo (R$ mil)"})
    if not tem_qtd:
        det = det.drop(columns=["Profundidade no topo (cotas)", "Profundidade no topo (R$ mil)"])
    det["ETF"] = pd.Categorical(det["ETF"], ordem, ordered=True)
    det = det.sort_values(["ETF", "Data"])

    notas = [
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
    ]
    if tem_qtd:
        notas.append("Profundidade no topo: média de (qtd_compra + qtd_venda) / 2 nas barras válidas, em cotas; "
                     "em R$ mil, cada lado vezes o seu preço.")
    notas += ["Fonte: LSEG Workspace. Dado bruto da LSEG: não redistribuir.",
              f"Gerado em {dt.datetime.now():%d/%m/%Y %H:%M}."]

    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        resumo.to_excel(w, sheet_name="Resumo", index=False)
        mensal.reset_index().to_excel(w, sheet_name="Mensal (bps)", index=False)
        diario.reset_index().to_excel(w, sheet_name="Diário (bps)", index=False)
        perfil.reset_index().to_excel(w, sheet_name="Perfil intradiário (bps)", index=False)
        det.to_excel(w, sheet_name="Diário detalhado", index=False)
        pd.DataFrame({"Notas": notas}).to_excel(w, sheet_name="Notas", index=False)
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
                    elif "cotas" in nome or "R$ mil" in nome:
                        c.number_format = "#,##0"
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


def relatorio(cfg, resumo, faltas, ld, arq_csv, arq_spread, executado):
    por_etf, sem_nenhum = faltas
    campos = cfg["campos"]

    falhas = []
    for x in resumo:
        if x["erro"]:
            falhas.append(f"{x['ticker']} ({x['ric']}): erro na chamada, ETF fora do arquivo. {x['erro']}")
        elif x["pregoes"] == 0:
            falhas.append(f"{x['ticker']} ({x['ric']}): sem nenhum dado no período. Conferir o RIC.")
        for c in x["campos_vazios"]:
            falhas.append(f"{x['ticker']}: o campo {campos[c]} (coluna {c}) veio vazio no período todo. "
                          "Conferir o nome do campo no Data Item Browser (DIB).")
    for t, dias in por_etf.items():
        falhas.append(f"{t}: {len(dias)} pregão(ões) sem dado, em que outros ETFs tiveram: {_datas(dias)}")

    linhas = [
        "Barras intraday de ETFs (LSEG Workspace)",
        f"Data de execução: {executado:%d/%m/%Y %H:%M}",
        f"Período: {cfg['data_inicio']:%d/%m/%Y} a {cfg['data_fim']:%d/%m/%Y}"
        + (" (rotina mensal: mês anterior)" if cfg["mensal"] else ""),
        f"Arquivo: {arq_csv.name}" + (f"; spread calculado: {arq_spread.name}" if arq_spread else ""),
        "",
        "0. Falhas",
    ]
    linhas += [f"   - {f}" for f in falhas] or ["   Nenhuma."]
    if sem_nenhum:
        linhas.append(f"   Dias úteis sem dado em nenhum ETF (provável feriado da B3; se não for, é falha): {_datas(sem_nenhum, 30)}")
    linhas += [
        "",
        "1. Campos usados na chamada (coluna do CSV = campo da API)",
    ]
    linhas += [f"   {c} = {campos[c]}" for c in COLUNAS]
    linhas += [
        f"   Biblioteca: {versao_biblioteca(ld)}, função get_history.",
        f"   universe = <RIC> (ex.: {resumo[0]['ric'] if resumo else 'BOVA11.SA'})",
        f"   fields = {campos_da_chamada(cfg)}",
        f"   interval = \"{cfg['intervalo_api']}\"",
        "   start/end = um mês calendário por chamada, em UTC (o mês é dividido se a resposta vier cortada).",
        "   Valor de cada barra: o último do intervalo, como a API devolve.",
        "",
        "2. Desde quando o histórico intradiário existe",
    ]
    for x in resumo:
        if x["erro"]:
            linhas.append(f"   {x['ticker']}: - (erro na chamada)")
        else:
            linhas.append(f"   {x['ticker']}: {x['historico_desde']} ({x['como']}); primeira barra no período: "
                          f"{x['primeira_barra_periodo']}")
    linhas += [
        "",
        "3. Tratamento",
        "   Horário: a API devolve em UTC; o CSV está no horário de Brasília.",
        f"   Uma linha a cada {cfg['intervalo_min']} minutos, de {cfg['janela_inicio']:%H:%M} a {cfg['janela_fim']:%H:%M} "
        "(inclusive), para cada ETF, só nos pregões com dado.",
        "   Barra sem atualização: repete o último valor do mesmo dia, campo a campo, nunca do dia anterior. Barras do",
        f"   próprio dia entre {cfg['abertura_continuo']:%H:%M} e {cfg['janela_inicio']:%H:%M} servem de ponto de partida; "
        "as do leilão de abertura, não.",
        "   Se o dia ainda não tinha oferta até aquele horário, a linha fica vazia (nunca zero).",
        "   Carimbo de hora: o da barra como a API devolve. Confirmar na documentação da LSEG se ele marca o",
        "   início ou o fim do intervalo (muda em 5 minutos a leitura de cada linha).",
        "   Nenhum outro filtro (livro cruzado, oferta zerada etc. vão como vieram).",
        "   CSV: UTF-8, vírgula como separador, ponto decimal, sem coluna de índice; preços com 2 casas e",
        "   quantidades (cotas) inteiras.",
        "",
        "4. Linhas por ETF",
    ]
    for x in resumo:
        if x["erro"]:
            linhas.append(f"   {x['ticker']}: erro, fora do arquivo")
        else:
            linhas.append(f"   {x['ticker']}: {x['pregoes']} pregões, {x['linhas']} linhas, "
                          f"{x['linhas_sem_compra_ou_venda']} linhas sem compra ou venda")
    linhas += [
        "",
        "5. Campo diário pronto de spread médio",
        "   Não verificado pelo script. Buscar 'spread' no Data Item Browser (DIB) do Workspace e conferir na",
        "   definição se é média do pregão ou foto do fechamento.",
        "",
        "Log da rodada:",
    ] + [f"   {l}" for l in LOG]
    return "\n".join(linhas) + "\n", len(falhas)


if __name__ == "__main__":
    main()
