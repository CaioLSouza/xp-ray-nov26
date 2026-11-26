"""Diagnóstico do dado tick a tick na LSEG (para as quantidades no topo do livro).

A barra de 5 minutos não traz BIDSIZE/ASKSIZE; o tick traz. Este teste responde três perguntas:
  1. até quantos dias para trás existe tick (de 1 a 120 dias);
  2. quanto pesa: linhas e segundos por hora de pregão;
  3. o carimbo da barra de 5 minutos marca o início ou o fim do intervalo (comparando com o tick).

Uso (Workspace aberto e logado):
    python diagnostico_tick.py                  # BOVA11.SA e LFTS11=SA
    python diagnostico_tick.py BOVA11.SA NCDI11=SA

Grava tudo em diagnostico_tick.txt, na pasta do script.
"""
import os
import sys
import time
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).resolve().parent
SAIDA = PASTA / "diagnostico_tick.txt"
FMT = "%Y-%m-%dT%H:%M:%S"
linhas = []


def out(msg=""):
    print(msg)
    linhas.append(str(msg))


def dia_util_atras(hoje, n):
    d = hoje - pd.Timedelta(days=n)
    while d.dayofweek >= 5:
        d -= pd.Timedelta(days=1)
    return d


def pede(ld, ric, campos, intervalo, ini, fim):
    t0 = time.time()
    try:
        df = ld.get_history(universe=ric, fields=campos, interval=intervalo, start=ini.strftime(FMT), end=fim.strftime(FMT))
        df = df if df is not None else pd.DataFrame()
        erro = ""
    except Exception as e:  # noqa: BLE001
        df, erro = pd.DataFrame(), f"{type(e).__name__}: {e}"
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[-1] for c in df.columns]
    return df, time.time() - t0, erro


def main():
    rics = sys.argv[1:] or ["BOVA11.SA", "LFTS11=SA"]
    try:
        import lseg.data as ld
    except ImportError:
        import refinitiv.data as ld
    chave = os.environ.get("LSEG_APP_KEY") or os.environ.get("REFINITIV_APP_KEY")
    ld.open_session(app_key=chave) if chave else ld.open_session()
    hoje = pd.Timestamp.now("UTC").normalize()
    try:
        for ric in rics:
            out(f"\n===== {ric} =====")
            out("1. Profundidade do tick (1 hora, 14h-15h UTC = 11h-12h em Brasília, BID/ASK/BIDSIZE/ASKSIZE)")
            for n in (1, 3, 7, 14, 21, 30, 45, 60, 90, 120):
                d = dia_util_atras(hoje, n)
                df, seg, erro = pede(ld, ric, ["BID", "ASK", "BIDSIZE", "ASKSIZE"], "tick",
                                     d + pd.Timedelta(hours=14), d + pd.Timedelta(hours=15))
                cheio = f", BIDSIZE preenchido={int(df['BIDSIZE'].notna().sum())}" if "BIDSIZE" in df.columns else ""
                out(f"   {n:>3} dias atrás ({d:%Y-%m-%d}): {len(df)} linhas em {seg:.1f}s{cheio}"
                    + (f" | ERRO {erro}" if erro else ""))

            out("2. Carimbo da barra de 5 minutos (ontem, 14h-15h UTC)")
            d = dia_util_atras(hoje, 1)
            ini, fim = d + pd.Timedelta(hours=14), d + pd.Timedelta(hours=15)
            barras, _, e1 = pede(ld, ric, ["BID", "ASK"], "5min", ini, fim)
            tick, _, e2 = pede(ld, ric, ["BID", "ASK"], "tick", ini - pd.Timedelta(minutes=10), fim + pd.Timedelta(minutes=10))
            if e1 or e2 or barras.empty or tick.empty:
                out(f"   sem como comparar: barras={len(barras)} tick={len(tick)} {e1} {e2}")
                continue
            barras.index = pd.to_datetime(barras.index)
            tick.index = pd.to_datetime(tick.index)
            tick = tick.dropna(subset=["BID", "ASK"]).sort_index()
            passo = pd.Timedelta(minutes=5)
            ok_fim = ok_ini = n = 0
            exemplos = []
            for t, b in barras.iterrows():
                antes_t = tick[tick.index < t]
                antes_t5 = tick[tick.index < t + passo]
                if antes_t.empty or antes_t5.empty:
                    continue
                n += 1
                fim_bate = (antes_t.iloc[-1][["BID", "ASK"]].values == b[["BID", "ASK"]].values).all()
                ini_bate = (antes_t5.iloc[-1][["BID", "ASK"]].values == b[["BID", "ASK"]].values).all()
                ok_fim += fim_bate
                ok_ini += ini_bate
                if len(exemplos) < 4:
                    exemplos.append(f"   barra {t:%H:%M} BID/ASK={b['BID']}/{b['ASK']} | último tick antes de {t:%H:%M}: "
                                    f"{antes_t.iloc[-1]['BID']}/{antes_t.iloc[-1]['ASK']} | antes de {t + passo:%H:%M}: "
                                    f"{antes_t5.iloc[-1]['BID']}/{antes_t5.iloc[-1]['ASK']}")
            out(f"   {n} barras comparadas. Bate com o último tick ANTES do carimbo (carimbo = fim): {ok_fim}. "
                f"Bate com o último tick antes do carimbo + 5 min (carimbo = início): {ok_ini}.")
            for x in exemplos:
                out(x)
    finally:
        try:
            ld.close_session()
        except Exception:  # noqa: BLE001
            pass
        SAIDA.write_text("\n".join(linhas), encoding="utf-8")
        print(f"\nGravado em {SAIDA}")


if __name__ == "__main__":
    main()
