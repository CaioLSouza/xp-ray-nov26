"""Extrai da LSEG (API lseg.data) as carteiras de ETFs que replicam índices MSCI com Brasil, para tirar delas o
FIF que o MSCI usa hoje, sem o arquivo do MSCI (Teo, 09/11/2026).

Ideia: um ETF que replica o índice carrega cada ação em proporção ao FIF x número de ações do MSCI. Com o peso de
cada papel na carteira, o preço e o número de ações, dá para recuperar o FIF. Vários ETFs dos mesmos índices
(MSCI Brazil, MSCI Brazil Small Cap, MSCI EM IMI e MSCI EM Small Cap) servem para conferir uns contra os outros.

Rodar num computador com o LSEG Workspace aberto e logado. Precisa de pandas, openpyxl e lseg-data.

    python extrai_lseg_etfs.py --teste    # 1 a 2 minutos: só o EWZ, carteira mais recente
    python extrai_lseg_etfs.py            # completo: 8 ETFs x 3 datas (uns 15 a 30 minutos)

Se cair no meio, rode de novo: ele pula o que já baixou. Resultado na pasta `saida_lseg_etfs`, ao lado do script;
suba na AWS o `etfs_msci_lseg.xlsx` (ou a pasta inteira).

O que ele faz:
1. Carteira de cada ETF em cada data (campo TR.FdInvestorPctPortfolio, o mesmo do cutoff_estimation.py do
   Antonio). Tenta também o número de ações de cada posição; se o campo não existir, segue só com o peso.
   Cada ETF tem RICs alternativos; usa o primeiro que responder.
2. Cadastro de todos os papéis das carteiras (país sede e da bolsa) e, para os brasileiros, preço em US$, número
   de ações, free float e valor de mercado na data de cada carteira e em 20/07/2026 (preço do pro forma de agosto).
3. Junta tudo num Excel: aba `carteiras` (uma linha por ETF, data e papel), aba `papeis` (dados dos papéis
   brasileiros por data) e aba `log` (o que deu certo e o que falhou).
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pandas as pd

try:
    import lseg.data as ld
except ImportError:
    import refinitiv.data as ld

# ETF -> (índice, RICs alternativos). Se algum RIC não responder, ele tenta o próximo.
FUNDOS = {
    "EWZ": ("MSCI Brazil 25/50", ["EWZ.P", "EWZ.K", "EWZ"]),
    "IBZL": ("MSCI Brazil", ["IBZL.L", "IDBZ.L", "IBZL"]),
    "XMBR": ("MSCI Brazil", ["XMBR.DE", "XMBR.L", "XMBR"]),
    "EWZS": ("MSCI Brazil Small Cap", ["EWZS.OQ", "EWZS.O", "EWZS"]),
    "EIMI": ("MSCI EM IMI", ["EIMI.L", "EMIM.L", "EIMI"]),
    "IEMG": ("MSCI EM IMI", ["IEMG.P", "IEMG.K", "IEMG"]),
    "EEMS": ("MSCI EM Small Cap", ["EEMS.P", "EEMS.K", "EEMS"]),
    "SPYM": ("MSCI EM Small Cap", ["SPYM.DE", "EMSM.L", "SPYM"]),
}
DATAS = ["recente", "2026-10-30", "2026-09-30"]
DATA_PROFORMA = "2026-07-20"
LOTE = 300
N = 5000

CAMPOS_CARTEIRA = [
    # (nomes das colunas, campos) - do mais completo para o mais simples
    (["calc_date", "ric", "nome_posicao", "pct", "acoes"],
     [f"TR.FdInvestorPctPortfolio(EndNum={N}).calcdate", f"TR.FdInvestorPctPortfolio(EndNum={N}).SecurityOwnedRIC",
      f"TR.FdInvestorPctPortfolio(EndNum={N}).instrument", f"TR.FdInvestorPctPortfolio(EndNum={N})",
      f"TR.FdInvestorSharesHeld(EndNum={N})"]),
    (["calc_date", "ric", "nome_posicao", "pct"],
     [f"TR.FdInvestorPctPortfolio(EndNum={N}).calcdate", f"TR.FdInvestorPctPortfolio(EndNum={N}).SecurityOwnedRIC",
      f"TR.FdInvestorPctPortfolio(EndNum={N}).instrument", f"TR.FdInvestorPctPortfolio(EndNum={N})"]),
    (["ric", "nome_posicao", "pct", "acoes"],
     [f"TR.FundHoldingRIC(EndNum={N})", f"TR.FundHoldingName(EndNum={N})",
      f"TR.FundPercentageOfFundAssets(EndNum={N})", f"TR.FundNumberOfShares(EndNum={N})"]),
]
ADRS_BR = {"NU.N", "XP.OQ", "XP.O", "STNE.OQ", "STNE.O", "PAGS.N", "AUGO.O", "AUGO.OQ", "INTR.OQ", "INTR.O", "JBS.N",
           "VINP.O", "AFYA.O", "PAX.O", "NEXA.K", "NEXA.N", "VTEX.N", "ZENV.O", "BRFS.N"}

log = []


def registra(etapa, item, ok, detalhe=""):
    log.append({"etapa": etapa, "item": item, "ok": ok, "detalhe": str(detalhe)[:200]})
    print(f"  {'ok ' if ok else 'ERRO'} {etapa} {item} {detalhe if not ok else ''}"[:220])


def carteira(fundo: str, rics: list[str], data: str) -> pd.DataFrame | None:
    par = {} if data == "recente" else {"SDate": data}
    for ric in rics:
        for nomes, campos in CAMPOS_CARTEIRA:
            try:
                d = ld.get_data(universe=ric, fields=campos, parameters=par)
                d = d.iloc[:, 1: 1 + len(nomes)].copy()
                d.columns = nomes
                d = d[d["ric"].notna() & (d["ric"].astype(str).str.len() > 0)]
                if d.empty:
                    raise ValueError("carteira vazia")
                d.insert(0, "fundo", fundo)
                d.insert(1, "ric_fundo", ric)
                d.insert(2, "data_pedida", data)
                if "calc_date" not in d.columns:
                    d["calc_date"] = data
                if "acoes" not in d.columns:
                    d["acoes"] = None
                registra("carteira", f"{fundo} {data}", True)
                print(f"     {fundo} ({ric}) {data}: {len(d)} posições, data da carteira {d['calc_date'].iloc[0]}, "
                      f"{'com' if d['acoes'].notna().any() else 'sem'} número de ações")
                return d
            except Exception as e:                      # noqa: BLE001 - tenta o próximo conjunto de campos / RIC
                erro = e
    registra("carteira", f"{fundo} {data}", False, erro)
    return None


_tam = {"lote": LOTE}


def pede(rics: list[str], campos: list[str], nomes: list[str], minimo: int = 10) -> list[pd.DataFrame]:
    try:
        d = ld.get_data(universe=rics, fields=campos)
        d = d.iloc[:, : len(nomes)].copy()
        d.columns = nomes
        return [d]
    except Exception as e:                              # noqa: BLE001
        if len(rics) <= minimo:
            registra("papeis", f"lote de {len(rics)} ({rics[0]}...)", False, e)
            return []
        if len(rics) >= _tam["lote"]:
            _tam["lote"] = max(len(rics) // 2, minimo)
        meio = len(rics) // 2
        return pede(rics[:meio], campos, nomes, minimo) + pede(rics[meio:], campos, nomes, minimo)


def em_lotes(rics, campos, nomes):
    partes, i = [], 0
    while i < len(rics):
        n = _tam["lote"]
        partes += pede(rics[i: i + n], campos, nomes)
        i += n
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=nomes)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--teste", action="store_true", help="só o EWZ, carteira mais recente")
    ap.add_argument("--saida", default=str(Path(__file__).resolve().parent / "saida_lseg_etfs"))
    a = ap.parse_args()
    fundos = {"EWZ": FUNDOS["EWZ"]} if a.teste else FUNDOS
    datas = ["recente"] if a.teste else DATAS
    pasta = Path(a.saida + ("_teste" if a.teste else ""))
    pasta.mkdir(parents=True, exist_ok=True)
    ld.open_session()
    try:
        # 1) carteiras
        print("1) Carteiras dos ETFs")
        cart = []
        for f, (indice, rics) in fundos.items():
            for data in datas:
                arq = pasta / f"carteira_{f}_{data}.csv"
                if arq.exists():
                    cart.append(pd.read_csv(arq))
                    continue
                d = carteira(f, rics, data)
                if d is not None:
                    d["indice"] = indice
                    d.to_csv(arq, index=False)
                    cart.append(d)
        if not cart:
            raise SystemExit("Nenhuma carteira voltou. Veja se o Workspace está aberto e logado e mande o print.")
        cart = pd.concat(cart, ignore_index=True)
        cart["calc_date"] = pd.to_datetime(cart["calc_date"], errors="coerce").dt.strftime("%Y-%m-%d")

        # 2) cadastro de todos os papéis e dados dos brasileiros
        print("2) Cadastro dos papéis")
        arq_cad = pasta / "cadastro.csv"
        rics = sorted(set(cart["ric"].dropna().astype(str)))
        if arq_cad.exists():
            cad = pd.read_csv(arq_cad)
            falta = [r for r in rics if r not in set(cad["ric"])]
        else:
            cad, falta = pd.DataFrame(), rics
        if falta:
            novo = em_lotes(falta, ["TR.CommonName", "TR.HQCountryCode", "TR.ExchangeCountryCode", "TR.ISIN"],
                            ["ric", "nome", "pais_sede", "pais_bolsa", "isin"])
            cad = pd.concat([cad, novo], ignore_index=True).drop_duplicates("ric")
            cad.to_csv(arq_cad, index=False)
        br = cad[(cad["pais_sede"] == "BR") | (cad["pais_bolsa"] == "BR") | cad["ric"].astype(str).str.endswith(".SA")
                 | cad["ric"].isin(ADRS_BR)]["ric"].tolist()
        print(f"   {len(rics)} papéis nas carteiras, {len(br)} brasileiros")
        datas_pap = sorted(set(cart["calc_date"].dropna()) | {DATA_PROFORMA})
        pap = []
        for d in datas_pap:
            arq = pasta / f"papeis_{d}.csv"
            if arq.exists():
                x = pd.read_csv(arq)
                falta = [r for r in br if r not in set(x["ric"])]
            else:
                x, falta = pd.DataFrame(), br
            if falta:
                campos = [f"TR.PriceClose(SDate={d},Curn=USD)", f"TR.SharesOutstanding(SDate={d})",
                          f"TR.FreeFloatPct(SDate={d})", f"TR.CompanyMarketCap(SDate={d},Curn=USD,Scale=6)",
                          f"TR.IssueMarketCap(SDate={d},Curn=USD,Scale=6)"]
                novo = em_lotes(falta, campos, ["ric", "preco_usd", "acoes_emitidas", "free_float_pct",
                                                "mcap_empresa_usd_mn", "mcap_papel_usd_mn"])
                x = pd.concat([x, novo], ignore_index=True)
                x.to_csv(arq, index=False)
            x["data"] = d
            pap.append(x)
            print(f"   {d}: {x['preco_usd'].notna().sum()} de {len(br)} papéis brasileiros com preço")
        pap = pd.concat(pap, ignore_index=True)
    finally:
        ld.close_session()

    # 3) Excel
    cart_br = cart.merge(cad, on="ric", how="left")
    resumo = (cart_br.assign(br=cart_br["ric"].isin(br))
              .groupby(["fundo", "indice", "data_pedida", "calc_date"])
              .agg(posicoes=("ric", "size"), posicoes_br=("br", "sum"), pct_total=("pct", "sum"),
                   com_acoes=("acoes", lambda s: s.notna().sum())).reset_index())
    saida = pasta / "etfs_msci_lseg.xlsx"
    with pd.ExcelWriter(saida) as xw:
        resumo.to_excel(xw, sheet_name="resumo", index=False)
        cart_br[cart_br["ric"].isin(br)].to_excel(xw, sheet_name="carteiras_br", index=False)
        pap.merge(cad, on="ric", how="left").to_excel(xw, sheet_name="papeis", index=False)
        pd.DataFrame(log).to_excel(xw, sheet_name="log", index=False)
    print("\nResumo:")
    print(resumo.to_string(index=False))
    print(f"\nGravado {saida}  <- suba este arquivo na AWS")


if __name__ == "__main__":
    main()
