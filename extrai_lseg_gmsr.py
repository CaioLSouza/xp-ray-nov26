"""Extrai da LSEG (API lseg.data) as ações dos países desenvolvidos com valor de mercado cheio e free float
em 14 datas, para refazer o GMSR do MSCI do jeito que o MSCI calcula (Teo, 09/11/2026).

Rodar num computador com o LSEG Workspace aberto e logado (é a mesma sessão que o `rd.open_session()` dos
códigos do Antonio usa). Precisa de pandas, openpyxl e lseg-data (ou refinitiv-data).

    python extrai_lseg_gmsr.py --teste     # 2 minutos: 3 países pequenos e 2 datas, só para ver se a API responde
    python extrai_lseg_gmsr.py             # completo: 24 países e 14 datas (uns 20 a 40 minutos)

Se cair no meio, é só rodar de novo: ele pula o que já baixou. O resultado fica na pasta `saida_lseg_gmsr`,
ao lado do script; suba na AWS o arquivo `gmsr_lseg_dm.xlsx` (ou a pasta inteira, se o Excel não abrir).

O que ele faz:
1. Universo: para cada país desenvolvido, roda o Screener (ações ativas, listagem principal, sede no país,
   valor de mercado da empresa >= US$ 500 mi hoje, perto do mínimo de tamanho do universo do MSCI) e
   deixa uma linha por empresa (Organization PermID), para não contar duas vezes quem tem mais de uma classe.
2. Para cada data, pede o valor de mercado da empresa em US$ mi, cheio e só o free float (ShType=FFL),
   os mesmos campos do `cutoff_estimation.py` do Antonio, em lotes de 300 empresas. Lote que der erro é
   dividido ao meio e tentado de novo.
3. Junta tudo num Excel com uma aba por medida (empresas nas linhas, datas nas colunas) e um resumo.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pandas as pd

try:
    import lseg.data as ld
except ImportError:                                   # versão antiga da biblioteca
    import refinitiv.data as ld

# 23 países desenvolvidos do MSCI + Luxemburgo (empresas com sede lá entram no índice de outro país)
PAISES = ["US", "CA", "GB", "IE", "FR", "DE", "NL", "BE", "LU", "CH", "SE", "NO", "DK", "FI", "ES", "PT", "IT",
          "AT", "IL", "JP", "AU", "NZ", "HK", "SG"]
# Datas de preço: as 4 reviews com GMSR publicado e data conhecida (para testar) e a janela de novembro/26
DATAS = ["2025-10-15", "2026-01-15", "2026-04-20", "2026-07-20",
         "2026-10-19", "2026-10-20", "2026-10-21", "2026-10-22", "2026-10-23",
         "2026-10-26", "2026-10-27", "2026-10-28", "2026-10-29", "2026-10-30"]
MCAP_MIN_USD_MN = 500
LOTE = 300


def tela(pais: str) -> str:
    return (f'SCREEN(U(IN(Equity(active,public,primary))/*UNV:Public*/), IN(TR.HQCountryCode,"{pais}"), '
            f'TR.CompanyMarketCap(Scale=6)>={MCAP_MIN_USD_MN}, CURN=USD)')


def universo(paises: list[str], pasta: Path) -> pd.DataFrame:
    arq = pasta / "universo.csv"
    if arq.exists():
        u = pd.read_csv(arq, dtype=str)
        print(f"Universo já baixado: {len(u)} empresas ({arq.name})")
        return u
    campos = ["TR.CommonName", "TR.OrganizationID", "TR.HQCountryCode", "TR.ExchangeCountryCode",
              "TR.CompanyMarketCapitalization(Scale=6,Curn=USD)"]
    nomes = ["RIC", "nome", "org_id", "pais_sede", "pais_bolsa", "mcap_hoje_usd_mn"]
    # plano B se algum campo de cadastro não existir na sua licença: só nome e valor de mercado
    campos_b, nomes_b = ["TR.CommonName", "TR.CompanyMarketCapitalization(Scale=6,Curn=USD)"], ["RIC", "nome", "mcap_hoje_usd_mn"]
    partes, falhas = [], []
    for p in paises:
        d = None
        for cs, ns in [(campos, nomes), (campos_b, nomes_b)]:
            try:
                d = ld.get_data(universe=tela(p), fields=cs)
                d = d.iloc[:, : len(ns)].copy()
                d.columns = ns
                break
            except Exception as e:                      # noqa: BLE001 - registra e tenta o plano B / próximo país
                erro = str(e)[:150]
                d = None
        if d is None:
            falhas.append(p)
            print(f"  {p}: ERRO {erro}")
            continue
        for c in nomes:
            if c not in d.columns:
                d[c] = p if c.startswith("pais") else None
        partes.append(d[nomes])
        print(f"  {p}: {len(d)} ações" + ("" if d["org_id"].notna().any() else " (sem PermID: plano B)"))
    if not partes:
        raise SystemExit("Nenhum país voltou do Screener. Veja se o Workspace está aberto e logado.")
    u = pd.concat(partes, ignore_index=True)
    u["mcap_hoje_usd_mn"] = pd.to_numeric(u["mcap_hoje_usd_mn"], errors="coerce")
    u["org_id"] = u["org_id"].astype(str).where(u["org_id"].notna(), u["RIC"])
    n0 = len(u)
    u = (u.sort_values(["mcap_hoje_usd_mn", "RIC"], ascending=[False, True])     # ordem estável entre rodadas
         .drop_duplicates("org_id").reset_index(drop=True))
    print(f"Universo: {n0} ações, {len(u)} empresas depois de juntar classes e recibos da mesma empresa")
    if falhas:
        print(f"ATENÇÃO, países que falharam (rode de novo depois): {', '.join(falhas)}")
    else:
        u.astype(str).to_csv(arq, index=False)         # só grava completo, para a próxima rodada não pular país
    return u


_tam = {"lote": LOTE}       # tamanho de lote que a API aceita; cai pela metade quando um lote cheio é recusado


def pede_lote(rics: list[str], data: str, minimo: int = 10) -> list[pd.DataFrame]:
    campos = [f"TR.CompanyMarketCapitalization(SDate={data},Scale=6,Curn=USD)",
              f"TR.CompanyMarketCapitalization(SDate={data},Scale=6,Curn=USD,ShType=FFL)"]
    for tentativa in range(1 if len(rics) > minimo else 2):   # lote grande que falha é dividido na hora
        try:
            d = ld.get_data(universe=rics, fields=campos)
            d = d.iloc[:, :3].copy()
            d.columns = ["RIC", "mcap_cheio_usd_mn", "mcap_ff_usd_mn"]
            return [d]
        except Exception as e:                          # noqa: BLE001
            erro = str(e)[:150]
            if len(rics) <= minimo:
                time.sleep(3)
    if len(rics) <= minimo:
        print(f"    lote de {len(rics)} sem resposta ({erro}); fica de fora: {', '.join(rics[:5])}...")
        return []
    if len(rics) >= _tam["lote"]:
        _tam["lote"] = max(len(rics) // 2, minimo)
        print(f"    a API recusou {len(rics)} de uma vez ({erro[:80]}); passo a pedir {_tam['lote']} por lote")
    meio = len(rics) // 2
    return pede_lote(rics[:meio], data, minimo) + pede_lote(rics[meio:], data, minimo)


def baixa_datas(u: pd.DataFrame, datas: list[str], pasta: Path, lote: int) -> None:
    _tam["lote"] = lote
    cols = ["RIC", "mcap_cheio_usd_mn", "mcap_ff_usd_mn"]
    for data in datas:
        arq = pasta / f"mcap_{data}.csv"
        feito = pd.read_csv(arq) if arq.exists() else pd.DataFrame(columns=cols)
        rics = [r for r in u["RIC"] if r not in set(feito["RIC"])]   # retomada: só o que falta
        if not rics:
            print(f"{data}: já baixado")
            continue
        t0, partes, i = time.time(), [], 0
        while i < len(rics):
            n = _tam["lote"]
            partes += pede_lote(rics[i: i + n], data)
            i += n
        d = pd.concat([feito] + partes, ignore_index=True)
        for c in cols[1:]:
            d[c] = pd.to_numeric(d[c], errors="coerce")
        d.to_csv(arq, index=False)
        ok = d.loc[d["RIC"].isin(set(u["RIC"])), "mcap_cheio_usd_mn"].notna().sum()
        print(f"{data}: {ok} de {len(u)} empresas com valor ({time.time() - t0:.0f}s)")


def consolida(u: pd.DataFrame, datas: list[str], pasta: Path) -> None:
    longo = []
    for data in datas:
        arq = pasta / f"mcap_{data}.csv"
        if arq.exists():
            longo.append(pd.read_csv(arq).assign(data=data))
    if not longo:
        print("Nada para consolidar.")
        return
    longo = pd.concat(longo, ignore_index=True)
    cheio = longo.pivot_table(index="RIC", columns="data", values="mcap_cheio_usd_mn", aggfunc="first")
    ff = longo.pivot_table(index="RIC", columns="data", values="mcap_ff_usd_mn", aggfunc="first")
    base = u.set_index("RIC")[["nome", "org_id", "pais_sede", "pais_bolsa", "mcap_hoje_usd_mn"]]
    resumo = []
    for data in cheio.columns:
        c = cheio[data].dropna().sort_values(ascending=False)
        f = ff[data].dropna()
        resumo.append({"data": data, "empresas_com_valor": len(c), "mcap_cheio_total_usd_bi": c.sum() / 1e3,
                       "mcap_ff_total_usd_bi": f.sum() / 1e3, "maior_usd_bi": c.iloc[0] / 1e3 if len(c) else None,
                       "posicao_500_usd_bi": c.iloc[499] / 1e3 if len(c) >= 500 else None,
                       "posicao_1000_usd_bi": c.iloc[999] / 1e3 if len(c) >= 1000 else None})
    resumo = pd.DataFrame(resumo)
    saida = pasta / "gmsr_lseg_dm.xlsx"
    with pd.ExcelWriter(saida) as xw:
        resumo.to_excel(xw, sheet_name="resumo", index=False)
        base.join(cheio, how="left").to_excel(xw, sheet_name="mcap_cheio_usd_mn")
        base[["nome", "pais_sede"]].join(ff, how="left").to_excel(xw, sheet_name="mcap_ff_usd_mn")
    print("\nResumo por data:")
    print(resumo.round(1).to_string(index=False))
    print(f"\nGravado {saida}  <- suba este arquivo na AWS")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--teste", action="store_true", help="3 países pequenos e 2 datas")
    ap.add_argument("--saida", default=str(Path(__file__).resolve().parent / "saida_lseg_gmsr"))
    ap.add_argument("--lote", type=int, default=LOTE)
    a = ap.parse_args()
    paises, datas = (["NZ", "SG", "IE"], ["2026-07-20", "2026-10-30"]) if a.teste else (PAISES, DATAS)
    pasta = Path(a.saida + ("_teste" if a.teste else ""))
    pasta.mkdir(parents=True, exist_ok=True)
    ld.open_session()
    try:
        u = universo(paises, pasta)
        baixa_datas(u, datas, pasta, a.lote)
    finally:
        ld.close_session()
    consolida(u, datas, pasta)


if __name__ == "__main__":
    main()
