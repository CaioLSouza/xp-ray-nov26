"""Carteiras mensais (dez/2024 a out/2026) dos ETFs que replicam o MSCI Brazil e o MSCI Brazil Small Cap, pela API
da LSEG, para medir historicamente o quanto cada ETF serve de cópia do índice (Teo, 09/11/2026).

A análise (fif_etfs_hist.py) compara as ações de cada posição com as ações do índice (número de ações x FIF) dos
pro formas do MSCI antes e depois de cada review (fev/25 a ago/26) e do arquivo de constituintes de 06/08/2026.

Rodar num computador com o LSEG Workspace aberto e logado (pandas, openpyxl e lseg-data):

    python extrai_lseg_etfs_hist.py --teste   # 1 minuto: só o IBZL em 2 datas
    python extrai_lseg_etfs_hist.py           # completo: 6 ETFs x 23 meses (uns 10 a 20 minutos)

Se cair no meio, rode de novo: ele pula o que já baixou. Suba na AWS o `etfs_msci_hist.xlsx` da pasta
`saida_lseg_etfs_hist`.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    import lseg.data as ld
except ImportError:
    import refinitiv.data as ld

# ETF -> (índice, identificadores: código da Lipper e da bolsa; usa o primeiro que responder)
FUNDOS = {
    "IBZL": ("MSCI Brazil", ["LP65013048", "IBZL.L", "IDBZ.L"]),                 # iShares MSCI Brazil UCITS (Irlanda)
    "ISHDE": ("MSCI Brazil", ["LP68515189"]),                                   # iShares MSCI Brazil UCITS (DE)
    "AMUNDI": ("MSCI Brazil", ["LP68539702"]),                                  # Amundi MSCI Brazil UCITS
    "XMBR": ("MSCI Brazil", ["LP65082851", "XMBR.DE"]),                         # Xtrackers MSCI Brazil UCITS
    "EWZ": ("MSCI Brazil 25/50", ["LP40056869", "EWZ.P"]),                      # iShares MSCI Brazil (EUA)
    "EWZS": ("MSCI Brazil Small Cap", ["LP40182878", "EWZS.OQ", "EWZS.O"]),     # iShares MSCI Brazil Small-Cap
}
DATAS = [d.strftime("%Y-%m-%d") for d in pd.date_range("2024-12-31", "2026-10-31", freq="BME")]
N = 5000
CAMPOS = [
    (["calc_date", "ric", "nome_posicao", "pct", "acoes"],
     [f"TR.FdInvestorPctPortfolio(EndNum={N}).calcdate", f"TR.FdInvestorPctPortfolio(EndNum={N}).SecurityOwnedRIC",
      f"TR.FdInvestorPctPortfolio(EndNum={N}).instrument", f"TR.FdInvestorPctPortfolio(EndNum={N})",
      f"TR.FdInvestorSharesHeld(EndNum={N})"]),
    (["calc_date", "ric", "nome_posicao", "pct"],
     [f"TR.FdInvestorPctPortfolio(EndNum={N}).calcdate", f"TR.FdInvestorPctPortfolio(EndNum={N}).SecurityOwnedRIC",
      f"TR.FdInvestorPctPortfolio(EndNum={N}).instrument", f"TR.FdInvestorPctPortfolio(EndNum={N})"]),
]
log = []


def carteira(fundo, ids, data):
    erro = None
    for ric in ids:
        for nomes, campos in CAMPOS:
            try:
                d = ld.get_data(universe=ric, fields=campos, parameters={"SDate": data})
                d = d.iloc[:, 1: 1 + len(nomes)].copy()
                d.columns = nomes
                d = d[d["ric"].notna() & (d["ric"].astype(str).str.len() > 0)]
                if d.empty:
                    raise ValueError("carteira vazia")
                if "acoes" not in d.columns:
                    d["acoes"] = None
                d.insert(0, "fundo", fundo)
                d.insert(1, "id_fundo", ric)
                d.insert(2, "data_pedida", data)
                log.append({"fundo": fundo, "data": data, "ok": True, "id": ric, "posicoes": len(d),
                            "calc_date": str(d["calc_date"].iloc[0]), "detalhe": ""})
                return d
            except Exception as e:                      # noqa: BLE001 - tenta o próximo campo / identificador
                erro = e
    log.append({"fundo": fundo, "data": data, "ok": False, "id": "", "posicoes": 0, "calc_date": "", "detalhe": str(erro)[:200]})
    return None


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--teste", action="store_true")
    ap.add_argument("--saida", default=str(Path(__file__).resolve().parent / "saida_lseg_etfs_hist"))
    a = ap.parse_args()
    fundos = {"IBZL": FUNDOS["IBZL"]} if a.teste else FUNDOS
    datas = ["2026-07-31", "2026-09-30"] if a.teste else DATAS
    pasta = Path(a.saida + ("_teste" if a.teste else ""))
    pasta.mkdir(parents=True, exist_ok=True)
    ld.open_session()
    partes = []
    try:
        for f, (indice, ids) in fundos.items():
            print(f"{f} ({indice})")
            for data in datas:
                arq = pasta / f"{f}_{data}.csv"
                if arq.exists():
                    partes.append(pd.read_csv(arq))
                    continue
                d = carteira(f, ids, data)
                if d is not None:
                    d["indice"] = indice
                    d.to_csv(arq, index=False)
                    partes.append(d)
                    print(f"   {data}: {len(d)} posições, carteira de {d['calc_date'].iloc[0]}, "
                          f"{'com' if d['acoes'].notna().any() else 'sem'} número de ações")
                else:
                    print(f"   {data}: ERRO {log[-1]['detalhe'][:120]}")
    finally:
        ld.close_session()
    if not partes:
        raise SystemExit("Nenhuma carteira voltou. Veja se o Workspace está aberto e logado e mande o print.")
    cart = pd.concat(partes, ignore_index=True)
    resumo = (cart.groupby(["fundo", "data_pedida", "calc_date"])
              .agg(posicoes=("ric", "size"), pct_total=("pct", "sum"), com_acoes=("acoes", lambda s: s.notna().sum()))
              .reset_index())
    saida = pasta / "etfs_msci_hist.xlsx"
    with pd.ExcelWriter(saida) as xw:
        resumo.to_excel(xw, sheet_name="resumo", index=False)
        cart.to_excel(xw, sheet_name="carteiras", index=False)
        pd.DataFrame(log).to_excel(xw, sheet_name="log", index=False)
    print(resumo.groupby("fundo").agg(datas=("data_pedida", "nunique"), datas_carteira=("calc_date", "nunique")).to_string())
    print(f"\nGravado {saida}  <- suba este arquivo na AWS")


if __name__ == "__main__":
    main()
