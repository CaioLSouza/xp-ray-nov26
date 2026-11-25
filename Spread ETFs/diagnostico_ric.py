"""Diagnóstico rápido: por que um ETF não traz dado na LSEG?

Uso (Workspace aberto e logado):
    python diagnostico_ric.py                      # BOVA11 (controle) + alguns de renda fixa
    python diagnostico_ric.py LFTS11 NCDI11        # os tickers que você quiser

Grava tudo em diagnostico_ric.txt, na pasta do script. Mande esse arquivo de volta.
"""
import os
import sys
import traceback
from pathlib import Path

import pandas as pd

PASTA = Path(__file__).resolve().parent
SAIDA = PASTA / "diagnostico_ric.txt"
CAMPOS = ["BID", "ASK", "BIDSIZE", "ASKSIZE", "TRDPRC_1"]
linhas = []


def out(msg=""):
    print(msg)
    linhas.append(str(msg))


def tenta(titulo, f):
    out(f"  -- {titulo}")
    try:
        r = f()
        if isinstance(r, pd.DataFrame):
            out(f"     linhas={len(r)} colunas={list(r.columns)}")
            if len(r):
                out("     preenchidos por coluna: " + ", ".join(f"{c}={int(r[c].notna().sum())}" for c in r.columns))
                out("     primeiras linhas:\n" + r.head(3).to_string())
                out("     últimas linhas:\n" + r.tail(3).to_string())
        else:
            out(f"     {r}")
    except Exception as e:  # noqa: BLE001
        out(f"     ERRO: {type(e).__name__}: {e}")
        out("     " + traceback.format_exc().strip().splitlines()[-1])


def main():
    tickers = [a.upper().replace(".SA", "") for a in sys.argv[1:]] or \
        ["BOVA11", "LFTS11", "NCDI11", "LTBX11", "AUPO11", "SELI11"]
    try:
        import lseg.data as ld
    except ImportError:
        import refinitiv.data as ld
    chave = os.environ.get("LSEG_APP_KEY") or os.environ.get("REFINITIV_APP_KEY")
    ld.open_session(app_key=chave) if chave else ld.open_session()
    fim = pd.Timestamp.now("UTC").floor("D")
    ini = fim - pd.Timedelta(days=7)
    fmt = "%Y-%m-%dT%H:%M:%S"
    out(f"Janela de teste: {ini:%Y-%m-%d} a {fim:%Y-%m-%d} (UTC)")
    try:
        for t in tickers:
            ric = f"{t}.SA"
            out(f"\n===== {t} ({ric}) =====")
            tenta("busca do RIC (discovery)", lambda: ld.discovery.search(
                query=t, select="RIC,DocumentTitle,ExchangeName,AssetState", top=5))
            tenta("nome e bolsa (get_data)", lambda: ld.get_data(ric, ["TR.CommonName", "TR.ExchangeName"]))
            tenta("5min, os 5 campos do script", lambda: ld.get_history(
                universe=ric, fields=CAMPOS, interval="5min", start=ini.strftime(fmt), end=fim.strftime(fmt)))
            tenta("5min, sem escolher campo (o que a API tem)", lambda: ld.get_history(
                universe=ric, interval="5min", start=ini.strftime(fmt), end=fim.strftime(fmt)))
            tenta("5min, só BID e ASK", lambda: ld.get_history(
                universe=ric, fields=["BID", "ASK"], interval="5min", start=ini.strftime(fmt), end=fim.strftime(fmt)))
            tenta("diário, BID/ASK/TRDPRC_1", lambda: ld.get_history(
                universe=ric, fields=["BID", "ASK", "TRDPRC_1"], interval="daily",
                start=(fim - pd.Timedelta(days=30)).strftime(fmt), end=fim.strftime(fmt)))
    finally:
        try:
            ld.close_session()
        except Exception:  # noqa: BLE001
            pass
        SAIDA.write_text("\n".join(linhas), encoding="utf-8")
        print(f"\nGravado em {SAIDA}")


if __name__ == "__main__":
    main()
