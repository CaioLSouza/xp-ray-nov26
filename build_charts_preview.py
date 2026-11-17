"""Monta a Charts_Preview (MSCI Brazil e MSCI Mexico) da review de novembro de 2026.

Pedido do Antonio (11/11/2026): uma aba para o MSCI Brazil e outra para o MSCI Mexico, com o fluxo
estimado de cada nome e o retorno contra um "control group" (índice setorial da B3 no Brasil, MEXBOL
no México), para os clientes que vão operar no dia do rebalanceamento (fechamento de 30/11/2026).

Fontes:
- Pesos, ADTV e fluxos: 202611 Rebalancing\\Actual\\202611 Rebalancing Flows.xlsx (o relatório de fluxos
  publicado em 11/11): Calc_BZ, Calc_Latam, Base_MSCI (pro forma do MSCI, preços de 20/10), JBS_migracao.
- Preços: Yahoo (B3 em fechamento ajustado por proventos; México em fechamento; NU, XP, AUGO e JBS pelos
  BDRs na B3) e site da B3 (fechamento oficial dos índices setoriais e do Ibovespa).

Uso:
  python build_charts_preview.py                 # último pregão com fechamento oficial da B3
  python build_charts_preview.py --last 2026-11-27
Saída: charts_preview\\Charts_Preview.xlsx (recalculada no Excel) e _data\\prices_<last>.csv
"""
from __future__ import annotations

import argparse
import base64
import json
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import yfinance as yf
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent
FLOWS = ROOT.parent / "202611 Rebalancing" / "Actual" / "202611 Rebalancing Flows.xlsx"
OUT = ROOT / "Charts_Preview.xlsx"
DATA = ROOT / "_data"

# ---------------------------------------------------------------- datas da review
CUTOFF = date(2026, 10, 20)      # data de preço do MSCI (conferida contra local_price do pro forma: BR e MX)
ANN = date(2026, 11, 11)         # anúncio (depois do fechamento)
DAY_AFTER = date(2026, 11, 12)   # primeiro pregão depois do anúncio
IMPL = date(2026, 11, 30)        # implementação no fechamento
EFFECTIVE = date(2026, 12, 1)
START = date(2026, 10, 14)       # início do histórico na aba Prices

# ---------------------------------------------------------------- control groups
# Índices setoriais da B3 que o Antonio indicou (prints na Flow Impact): IEE, IFNC, ICON, INDX, IMAT, IMOB.
# Cada nome vai para o índice que tem o papel na carteira teórica da B3 (16/11/2026); quem não está em
# nenhum vai pelo grupo do setor. Saneamento (SBSP3) vai para o UTIL; petróleo e telecom, que não têm
# índice setorial na B3, vão para o Ibovespa.
CONTROL_BR = {
    "IEEX": ["AXIA3", "AXIA7", "CMIG4", "CPLE3", "CPFE3", "ENGI11", "ENEV3", "EGIE3", "EQTL3"],
    "UTIL": ["SBSP3"],
    "IFNC": ["B3SA3", "BBDC3", "BBDC4", "BBAS3", "BPAC11", "BBSE3", "CXSE3", "ITUB3", "ITUB4", "ITSA4",
             "PSSA3", "NU", "XP"],
    "ICON": ["ABEV3", "RENT3", "MBRF3", "RADL3", "RDOR3"],
    # JBS é um ativo em dólar (listada na NYSE, receita global): contra o ICON (consumo doméstico, -27% com a
    # eleição) o relativo desde o corte dá +34 p.p. sem relação com o MSCI. Fica contra o par mais próximo.
    "MBRF3": ["JBS"],
    "IMAT": ["GGBR4", "KLBN11", "SUZB3", "VALE3", "AUGO"],
    "INDX": ["EMBJ3", "WEGE3", "MOTV3", "RAIL3"],
    "IBOV": ["PETR3", "PETR4", "PRIO3", "UGPA3", "VBBR3", "VIVT3", "TIMS3"],
}
CONTROL_INFO = [  # código B3, nome, ticker Bloomberg (prints do Antonio), quem usa
    ("IEEX", "Electric Utilities Index (IEE)", "IBOVIEE Index", "Utilities (electricity)"),
    ("IFNC", "Financials Index", "IFNCBV Index", "Banks, financials non-banks, NU, XP"),
    ("ICON", "Consumer Index", "ICONBV Index", "Food & beverages, retail, health care, car rental"),
    ("INDX", "Industrials Index", "INDX Index", "Capital goods, transportation"),
    ("IMAT", "Basic Materials Index", "IMAT Index", "Metals & mining, pulp & paper, AUGO"),
    ("IMOB", "Real Estate Index", "IMOBBV Index", "No MSCI Brazil member this cycle"),
    ("UTIL", "Public Utilities Index", "UTILBV Index (to confirm)", "Sanitation (SBSP3); not in Antonio's list"),
    ("IBOV", "Ibovespa", "IBOV Index", "Oil & gas and telecom (no B3 sector index)"),
    ("MBRF3", "MBRF (single stock)", "MBRF3 BZ Equity", "JBS only: closest listed protein exporter. ICON is "
     "domestic consumer and fell 27% after the election, while JBS trades as a US$ asset"),
    ("MEXBOL", "S&P/BMV IPC", "MEXBOL Index", "All MSCI Mexico members"),
]
B3_INDICES = ["IEEX", "IFNC", "ICON", "INDX", "IMAT", "IMOB", "UTIL", "IBOV"]

# Listados nos EUA: retorno pelo BDR na B3 (BRL, mesmo horário de fechamento do índice de controle)
BDR = {"NU": "ROXO34", "XP": "XPBR31", "AUGO": "AURA33", "JBS": "JBSS32"}

MX_NAMES = {
    "AMXB": "America Movil", "AC*": "Arca Continental", "CEMEXCPO": "Cemex", "KOFUBL": "Coca-Cola Femsa",
    "FEMSAUBD": "Femsa", "FUNO11": "Fibra Uno", "GRUMAB": "Gruma", "OMAB": "OMA (Aeroportuario Centro Norte)",
    "GAPB": "GAP (Aeroportuario del Pacifico)", "ASURB": "ASUR (Aeroportuario del Sureste)",
    "BIMBOA": "Grupo Bimbo", "GCARSOA1": "Grupo Carso", "GFNORTEO": "Banorte", "GFINBURO": "Inbursa",
    "GMEXICOB": "Grupo Mexico", "PE&OLES*": "Penoles", "KIMBERA": "Kimberly-Clark de Mexico",
    "FIBRAPL": "Fibra Prologis", "PINFRA*": "Pinfra", "SIGMAFA": "Sigma Foods", "WALMEX*": "Walmex",
}
MX_YAHOO = {"FIBRAPL": "FIBRAPL14.MX"}
SECTOR_SHORT = {
    "Equity Real Estate Investment Trusts (REITs)": "REITs",
    "Consumer Staples Distribution & Retail": "Consumer Staples Retail",
    "Telecommunication Services": "Telecom Services",
    "Food Beverage & Tobacco": "Food, Beverage & Tobacco",
}

H = {"User-Agent": "Mozilla/5.0"}


# ================================================================ dados do relatório de fluxos
def _calc(sheet: str) -> pd.DataFrame:
    raw = pd.read_excel(FLOWS, sheet_name=sheet, header=None)
    hrow = raw.index[raw.apply(lambda r: r.astype(str).str.strip().eq("Ticker").any(), axis=1)][0]
    c0 = int(np.where(raw.loc[hrow].astype(str).str.strip() == "Ticker")[0][0])
    names = ["ticker", "name", "sector", "adtv", "w_cur", "w_new", "delta_bps",
             "p_country", "p_latam", "p_em", "p_acwi", "p_total",
             "a_country", "a_latam", "a_em", "a_acwi", "a_total", "total", "imp_min", "imp_max"]
    df = raw.loc[hrow + 1:, c0:c0 + len(names) - 1].copy()
    df.columns = names
    # Calc_BZ: ticker Bloomberg 10 colunas antes do Ticker; Calc_Latam: país 11 antes, ticker Bloomberg 10 antes
    df["bbg"] = raw.loc[hrow + 1:, c0 - 10]
    df["cty"] = raw.loc[hrow + 1:, c0 - 11]
    df = df[df["ticker"].notna()].reset_index(drop=True)
    for c in names[3:]:
        df[c] = pd.to_numeric(df[c])
    return df


def base_msci() -> pd.DataFrame:
    return pd.read_excel(FLOWS, sheet_name="Base_MSCI")


def jbs_net_passive() -> float:
    raw = pd.read_excel(FLOWS, sheet_name="JBS_migracao", header=None)
    row = raw[raw[0].astype(str).str.startswith("L") & raw[0].astype(str).str.contains("quido")]
    return float(row.iloc[0, 1])


def aum_country() -> dict:
    raw = pd.read_excel(FLOWS, sheet_name="Base_Fundos", header=None)
    out = {}
    for _, r in raw.iterrows():
        if r[0] in ("Brazil", "Latam", "EM", "ACWI"):
            out[r[0]] = (float(r[1]), float(r[2]))  # (ativo, passivo)
    return out


# export da Lipper do Caio de 16/11 (bucket: "MSCI Brazil AUM (3).xlsx", apesar do nome é o México; inclui fundos inativos)
LIPPER_MX = ROOT.parent / "recebido_1611" / "lipper_mexico_v3" / "MSCI_Mexico_AUM_lipper_v3.xlsx"
# Fundos passivos que seguem o MSCI Mexico e não vieram na tela da Lipper (FSCREEN de 16/11/2026)
EXTRA_MX = {"iShares MSCI Mexico Capped UCITS ETF": ("CMX1.L", 116.476)}  # US$ mi, Yahoo totalAssets 16/11


def aum_mexico() -> tuple[float, str]:
    """Passivo que segue o MSCI Mexico: export da Lipper (FSCREEN, benchmark contém 'MSCI Mexico', qualquer
    asset universe) + os fundos que a tela não trouxe. Mesmo critério do fundos_aum_lipper.py: Index Tracking
    e Management Approach = Passive."""
    d = pd.read_excel(LIPPER_MX, header=0, skiprows=[1])
    d["TNA"] = pd.to_numeric(d["Fund TNA (Mil) (US Dollar).1"], errors="coerce")
    d["data"] = pd.to_datetime(d["Fund TNA (Mil) (US Dollar)"], errors="coerce")
    vivo = d["TNA"].gt(0) & d["data"].ge("2026-06-01")  # fundo fechado ou sem patrimônio recente fica fora
    pas = d[vivo & d["Management Approach"].eq("Passive") & d["Index Tracking"].fillna(0).astype(float).eq(1)]
    atv = d[vivo & d["Management Approach"].eq("Active")]
    total = float(pas["TNA"].sum()) + sum(v for _, v in EXTRA_MX.values())
    nomes = "; ".join(f"{n.replace(' Fund', '')} {v:,.0f}" for n, v in
                      pas.sort_values("TNA", ascending=False)[["Asset Name", "TNA"]].itertuples(index=False))
    extra = "; ".join(f"{n} {v:,.0f} (Yahoo {t}, not in the Lipper screen)" for n, (t, v) in EXTRA_MX.items())
    src = (f"Lipper (FSCREEN, TNA of {d['data'].max():%d-%b-%Y}), passive funds benchmarked to MSCI Mexico: {nomes}; "
           f"plus {extra}. Active funds with the same benchmark: US$ {atv['TNA'].sum():,.0f} mn "
           f"({len(atv)} funds, mostly Mexico Fund Inc). {int((~vivo).sum())} funds in the screen without TNA since Jun-2026 "
           "(closed or merged) are left out.")
    return total, src


def change_label(row) -> str:
    fam, sec = row.get("family_change_status"), row.get("security_change_status")
    if fam == "O" or pd.isna(row.get("pro_forma_index_market_cap")) or row.get("pro_forma_index_market_cap") == 0:
        return "Deletion"
    parts = []
    if isinstance(sec, str) and "F" in sec:
        parts.append(f"FIF {row['current_fif']:.1%} > {row['pro_forma_fif']:.1%}")
    if isinstance(sec, str) and "S" in sec:
        ch = row["pro_forma_number_of_shares"] / row["current_number_of_shares"] - 1
        parts.append(f"Shares {ch:+.1%}" if abs(ch) >= 0.0005 else "Shares (<0.1%)")
    return "; ".join(parts) if parts else "-"


LATAM = ["BR", "CL", "CO", "MX", "PE"]


def em_in_acwi() -> float:
    """Peso do EM no ACWI usado no relatório de fluxos (Calc_Ctry_Weights, 'ACWI weight')."""
    raw = pd.read_excel(FLOWS, sheet_name="Calc_Ctry_Weights", header=None)
    pos = np.argwhere(raw.astype(str).values == "ACWI weight")[0]
    return float(raw.iat[pos[0], pos[1] + 1])


def index_deltas(bm: pd.DataFrame) -> pd.DataFrame:
    """Método por índice: variação do peso de cada papel no MSCI EM Latin America e no MSCI EM (bps), pelo
    market cap no índice antes e depois da revisão. O fundo regional ou global vende o que sai de um papel e
    reinveste no índice inteiro, não só no país (o relatório de 11/11 reinvestia no país)."""
    cur = bm["current_index_market_cap"].fillna(0)
    pf = bm["pro_forma_index_market_cap"].fillna(0)
    la = bm["cty_code"].isin(LATAM)
    out = pd.DataFrame({"security_ticker": bm["security_ticker"]})
    out["d_em"] = (pf / pf.sum() - cur / cur.sum()) * 1e4
    out["d_la"] = np.where(la, (pf / pf[la].sum() - cur / cur[la].sum()) * 1e4, np.nan)
    return out.set_index("security_ticker")


def brazil_table(bm: pd.DataFrame) -> pd.DataFrame:
    bz = _calc("Calc_BZ")
    bmi = bm.set_index("security_ticker")
    bz["change"] = [change_label(bmi.loc[b]) if b in bmi.index else "-" for b in bz["bbg"]]
    bz.loc[bz["ticker"] == "JBS", "change"] = "Deletion (reclassified as US)"
    dl = index_deltas(bm)
    assert bz["bbg"].isin(dl.index).all(), bz.loc[~bz["bbg"].isin(dl.index), "bbg"].tolist()
    bz["d_la"] = bz["bbg"].map(dl["d_la"])
    bz["d_em"] = bz["bbg"].map(dl["d_em"])
    # JBS: além da saída do Standard, vendem os fundos de EM IMI e ACWI ex USA IMI e compram os de small cap
    # americanos (aba JBS_migracao); entra como "outro passivo", o líquido além do Standard (−55 mi)
    j = bz["ticker"] == "JBS"
    bz["p_other"] = 0.0
    bz.loc[j, "p_other"] = jbs_net_passive() - bz.loc[j, "p_total"]
    ctrl = {t: c for c, ts in CONTROL_BR.items() for t in ts}
    missing = set(bz["ticker"]) - set(ctrl)
    assert not missing, f"sem control group: {missing}"
    bz["control"] = bz["ticker"].map(ctrl)
    bz["series"] = bz["ticker"].map(lambda t: BDR.get(t, t))
    return bz


def mexico_table(bm: pd.DataFrame) -> pd.DataFrame:
    lt = _calc("Calc_Latam")
    lt = lt[lt["cty"] == "MX"].copy()
    mx = bm[bm["cty_code"] == "MX"].copy()
    mx["ticker"] = mx["security_ticker"].str.replace(" MF Equity", "", regex=False)
    mx["w_cur"] = mx["current_index_market_cap"] / mx["current_index_market_cap"].sum()
    mx["w_new"] = mx["pro_forma_index_market_cap"] / mx["pro_forma_index_market_cap"].sum()
    mx["change"] = mx.apply(change_label, axis=1)
    lt["key"] = lt["ticker"].astype(str)
    m = mx.merge(lt[["key", "sector", "adtv"]], left_on="ticker", right_on="key", how="left", validate="1:1")
    assert m["adtv"].notna().all(), m.loc[m["adtv"].isna(), "ticker"].tolist()
    dl = index_deltas(bm)
    m["d_la"] = m["security_ticker"].map(dl["d_la"])
    m["d_em"] = m["security_ticker"].map(dl["d_em"])
    m["p_other"] = 0.0
    m["name"] = m["ticker"].map(MX_NAMES)
    m["sector"] = m["sector"].map(lambda s: SECTOR_SHORT.get(s, s))
    m["control"] = "MEXBOL"
    m["series"] = m["ticker"].str.replace("*", "", regex=False)
    m["yahoo"] = m["ticker"].map(lambda t: MX_YAHOO.get(t, t.replace("*", "") + ".MX"))
    m["delta_bps"] = (m["w_new"] - m["w_cur"]) * 1e4
    return m


# ================================================================ preços
def b3_index(code: str, years) -> pd.Series:
    out = {}
    for y in years:
        p = base64.b64encode(json.dumps({"language": "pt-br", "index": code, "year": str(y)}).encode()).decode()
        r = requests.get("https://sistemaswebb3-listados.b3.com.br/indexStatisticsProxy/IndexCall/"
                         f"GetPortfolioDay/{p}", headers=H, timeout=30)
        for row in r.json().get("results", []):
            for mth in range(1, 13):
                v = row.get(f"rateValue{mth}")
                if v:
                    out[pd.Timestamp(y, mth, int(row["day"]))] = float(v.replace(".", "").replace(",", "."))
    return pd.Series(out, name=code).sort_index()


def yahoo(symbols: list[str], field: str, end: date) -> pd.DataFrame:
    d = yf.download(symbols, start=str(START), end=str(end + pd.Timedelta(days=1)),
                    auto_adjust=False, progress=False)[field]
    if isinstance(d, pd.Series):
        d = d.to_frame(symbols[0])
    d.index = pd.to_datetime(d.index).tz_localize(None).normalize()
    vazios = [s for s in symbols if s not in d.columns or d[s].dropna().empty]
    if vazios:
        raise RuntimeError(f"Yahoo sem preço para: {vazios}")
    return d


def prices(bz: pd.DataFrame, mx: pd.DataFrame, last: date | None):
    idx = pd.concat([b3_index(c, [2026]) for c in B3_INDICES], axis=1)
    if last is None:
        last = idx["IBOV"].dropna().index.max().date()
    idx = idx.loc[str(START):str(last)]

    br_series = list(dict.fromkeys(bz["series"]))
    br = yahoo([f"{s}.SA" for s in br_series], "Adj Close", last)
    br.columns = [c.replace(".SA", "") for c in br.columns]
    mxp = yahoo(mx["yahoo"].tolist() + ["^MXX"], "Close", last)
    ymap = dict(zip(mx["yahoo"], mx["series"]))
    ymap["^MXX"] = "MEXBOL"
    mxp.columns = [ymap[c] for c in mxp.columns]

    br = br[br_series]
    mxp = mxp[list(mx["series"]) + ["MEXBOL"]]
    px = pd.concat([br, idx, mxp], axis=1).sort_index().loc[str(START):str(last)]
    px = px[px.index.dayofweek < 5]
    filled = px.isna() & px.ffill().notna()
    px = px.ffill()

    meta = {}
    for s in br_series:
        src = next((k for k, v in BDR.items() if v == s), None)
        meta[s] = (f"{s} BDR of {src} (BRL)" if src else f"{s} (BRL)", "Yahoo, close adjusted for dividends")
    for c in B3_INDICES:
        meta[c] = (dict((a, b) for a, b, *_ in CONTROL_INFO)[c], "B3, official close")
    for s, t in zip(mx["series"], mx["ticker"]):
        meta[s] = (f"{t} (MXN)", "Yahoo, close")
    meta["MEXBOL"] = ("S&P/BMV IPC (MXN)", "Yahoo (^MXX), close")
    return px, filled, meta, last


# ================================================================ planilha
NAVY, LIGHT, YELLOW, GREY, RED = "1F2F44", "8EB3DF", "F9C113", "7E7E7E", "C00000"
F_BODY = Font(name="Roboto Light", size=9)
F_BOLD = Font(name="Roboto", size=9, bold=True)
F_HEAD = Font(name="Roboto", size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name="Roboto", size=12, bold=True, color=NAVY)
F_SUB = Font(name="Roboto Light", size=9, color=GREY)
F_NOTE = Font(name="Roboto Light", size=8, color=GREY)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_ZEBRA = PatternFill("solid", fgColor="F2F2F2")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
WHITE = Side(style="thin", color="FFFFFF")
B_HEAD = Border(left=WHITE, right=WHITE, top=WHITE, bottom=WHITE)
B_TOP = Border(top=Side(style="thin", color="000000"))
DATEF = "[$-409]dd-mmm-yy"

PX_LAST_COL = "DZ"  # folga de colunas na aba Prices
PX_LAST_ROW = 2000


def px_ref(date_name: str, code_cell: str) -> str:
    return (f"INDEX(Prices!$A$1:${PX_LAST_COL}${PX_LAST_ROW},MATCH({date_name},PxDates,1)+3,"
            f"MATCH({code_cell},Prices!$A$1:${PX_LAST_COL}$1,0))")


def ret(d1: str, d0: str, code_cell: str) -> str:
    return f'=IFERROR({px_ref(d1, code_cell)}/{px_ref(d0, code_cell)}-1,"n.a.")'


COLS = [  # chave, cabeçalho, grupo, largura, formato
    ("ticker", "Ticker", "", 10, None),
    ("name", "Name", "", 28, None),
    ("sector", "Sector", "", 20, None),
    ("change", "Index change", "", 22, None),
    ("w_cur", "Current", "Weight in the index", 8.5, "0.00%"),
    ("w_new", "New", "Weight in the index", 8.5, "0.00%"),
    ("dw", "Change (bps)", "Weight in the index", 8.5, "0.0"),
    ("adtv", "3M ADTV (US$ mn)", "Weight in the index", 9.5, "#,##0.0"),
    ("pc", "Passive: country funds", "Estimated flows (US$ mn)", 10, "#,##0.0"),
    ("pr", "Passive: regional & global funds", "Estimated flows (US$ mn)", 11, "#,##0.0"),
    ("pt", "Passive total", "Estimated flows (US$ mn)", 9.5, "#,##0.0"),
    ("tt", "Total (passive + active)", "Estimated flows (US$ mn)", 10, "#,##0.0"),
    ("ip", "Passive", "Impact (days of ADTV)", 8, "0.00"),
    ("it", "Total", "Impact (days of ADTV)", 8, "0.00"),
    ("s_cut", "Since cutoff", "Stock return", 9, "0.0%"),
    ("s_day", "Day after", "Stock return", 9, "0.0%"),
    ("s_ann", "Since announcement", "Stock return", 12.5, "0.0%"),
    ("control", "Control group", "Control group return", 9, None),
    ("c_cut", "Since cutoff", "Control group return", 9, "0.0%"),
    ("c_day", "Day after", "Control group return", 9, "0.0%"),
    ("c_ann", "Since announcement", "Control group return", 12.5, "0.0%"),
    ("r_cut", "Since cutoff", "Relative return (stock - control, p.p.)", 9, "0.0%"),
    ("r_day", "Day after", "Relative return (stock - control, p.p.)", 9, "0.0%"),
    ("r_ann", "Since announcement", "Relative return (stock - control, p.p.)", 12.5, "0.0%"),
    ("d_la", "MSCI EM Latin America", "Change in weight (bps) and other flows", 10, "0.00"),
    ("d_em", "MSCI EM", "Change in weight (bps) and other flows", 9, "0.000"),
    ("p_other", "Other passive (US$ mn)", "Change in weight (bps) and other flows", 10, "#,##0.0"),
    ("series", "Price series (Prices tab)", "Lookup key", 11, None),
]
COLIDX = {k: i + 1 for i, (k, *_) in enumerate(COLS)}


def L(key: str) -> str:
    return get_column_letter(COLIDX[key])


def write_market_tab(wb, title: str, df: pd.DataFrame, aum_name: str, subtitle: str, notes: list[str]):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws["A1"] = f"{title}: November 2026 Index Review, estimated flows and price action"
    ws["A1"].font = F_TITLE
    ws["A2"] = subtitle
    ws["A2"].font = F_SUB

    for i, (lab, name) in enumerate([("MSCI price cutoff", "CutoffDate"), ("Announcement (after close)", "AnnDate"),
                                     ("Day after", "DayAfterDate"), ("Last price", "LastDate"),
                                     ("Implementation (close)", "ImplDate")]):
        r = 3 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(r, 1, lab).font = F_BOLD
        c = ws.cell(r, 3, f"={name}")
        c.number_format, c.font, c.alignment = DATEF, F_BODY, Alignment(horizontal="left")
    ws.merge_cells("E3:R7")
    ws["E3"] = ("How to read: the control group proxies what the stock would have done without the index flows "
                "(B3 sector index for Brazil, MEXBOL for Mexico). Relative return = stock return minus control "
                "group return, in percentage points. A name with inflows that already outperformed its control "
                "may have the flow partly priced in; one that lagged may not. Dates are inputs on the Basis tab; "
                "prices are on the Prices tab." + (
                    " This cycle, the window since the cutoff includes the 26-Oct post-election sell-off "
                    "(Ibovespa -13.9%, BRL -8%): read the since-announcement and day-after columns for the "
                    "rebalance effect." if title == "MSCI Brazil" else ""))
    ws["E3"].font, ws["E3"].alignment = F_SUB, Alignment(wrap_text=True, vertical="top")

    GR, HR, R0 = 9, 10, 11
    # cabeçalho em dois níveis
    groups = {}
    for k, head, grp, width, fmt in COLS:
        j = COLIDX[k]
        ws.column_dimensions[get_column_letter(j)].width = width
        c = ws.cell(HR, j, head)
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, B_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        groups.setdefault(grp, []).append(j)
    for grp, js in groups.items():
        a, b = min(js), max(js)
        for j in range(a, b + 1):
            c = ws.cell(GR, j)
            c.fill, c.border = FILL_HEAD, B_HEAD
        if grp:
            if b > a:
                ws.merge_cells(start_row=GR, start_column=a, end_row=GR, end_column=b)
            c = ws.cell(GR, a, grp)
            c.font, c.alignment = F_HEAD, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[HR].height = 36
    ws.row_dimensions[GR].height = 26

    for i, row in enumerate(df.itertuples(index=False)):
        r = R0 + i
        rr = str(r)
        v = {
            "ticker": row.ticker, "name": row.name, "sector": row.sector, "change": row.change,
            "w_cur": float(row.w_cur), "w_new": float(row.w_new),
            "dw": f"=({L('w_new')}{rr}-{L('w_cur')}{rr})*10000",
            "adtv": float(row.adtv),
            "pc": f"={L('dw')}{rr}/10000*{aum_name}",
            "pr": (f"=({L('d_la')}{rr}*AUM_LATAM+{L('d_em')}{rr}*(AUM_EM+EM_IN_ACWI*AUM_ACWI))/10000"
                   f"+{L('p_other')}{rr}"),
            "d_la": float(row.d_la), "d_em": float(row.d_em), "p_other": float(row.p_other),
            "pt": f"={L('pc')}{rr}+{L('pr')}{rr}",
            "tt": f"={L('pt')}{rr}*(1+ActiveRatio)",
            "ip": f"={L('pt')}{rr}/{L('adtv')}{rr}",
            "it": f"={L('tt')}{rr}/{L('adtv')}{rr}",
            "s_cut": ret("LastDate", "CutoffDate", f"${L('series')}{rr}"),
            "s_day": ret("DayAfterDate", "AnnDate", f"${L('series')}{rr}"),
            "s_ann": ret("LastDate", "AnnDate", f"${L('series')}{rr}"),
            "control": row.control,
            "c_cut": ret("LastDate", "CutoffDate", f"${L('control')}{rr}"),
            "c_day": ret("DayAfterDate", "AnnDate", f"${L('control')}{rr}"),
            "c_ann": ret("LastDate", "AnnDate", f"${L('control')}{rr}"),
            "series": row.series,
        }
        for k_r, k_s, k_c in [("r_cut", "s_cut", "c_cut"), ("r_day", "s_day", "c_day"), ("r_ann", "s_ann", "c_ann")]:
            v[k_r] = (f'=IF(AND(ISNUMBER({L(k_s)}{rr}),ISNUMBER({L(k_c)}{rr})),'
                      f'{L(k_s)}{rr}-{L(k_c)}{rr},"n.a.")')
        for k, head, grp, width, fmt in COLS:
            c = ws.cell(r, COLIDX[k], v[k])
            c.font = F_BOLD if k in ("ticker", "r_cut", "r_day", "r_ann") else F_BODY
            if k == "change" and v[k] != "-":
                c.font = Font(name="Roboto", size=9, bold=True, color=NAVY)
            if fmt:
                c.number_format = fmt
            c.alignment = Alignment(horizontal="left" if k in ("ticker", "name", "sector", "change", "series")
                                    else "center" if k == "control" else "right")
            if i % 2 == 1:
                c.fill = FILL_ZEBRA
        ws.cell(r, COLIDX["series"]).font = F_NOTE
    last_r = R0 + len(df) - 1

    # negativos em vermelho XP
    for k in ("dw", "pc", "pr", "pt", "tt", "ip", "it", "s_cut", "s_day", "s_ann", "c_cut", "c_day", "c_ann",
              "r_cut", "r_day", "r_ann", "d_la", "d_em", "p_other"):
        rng = f"{L(k)}{R0}:{L(k)}{last_r}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED)))

    # fio preto no pé da tabela e notas
    for j in range(1, len(COLS) + 1):
        ws.cell(last_r + 1, j).border = B_TOP
    for i, n in enumerate(notes):
        c = ws.cell(last_r + 2 + i, 1, n)
        c.font = F_NOTE
    ws.freeze_panes = ws.cell(R0, 3)
    ws.auto_filter.ref = f"A{HR}:{L('series')}{last_r}"
    ws.sheet_view.zoomScale = 100
    ws.row_dimensions[3].height = 14
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{GR}:{HR}"
    return ws


def write_prices(wb, px: pd.DataFrame, filled: pd.DataFrame, meta: dict):
    ws = wb.create_sheet("Prices")
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A2"], ws["A3"] = "Series", "Description", "Source"
    for c in ("A1", "A2", "A3"):
        ws[c].font, ws[c].fill = F_HEAD, FILL_HEAD
    ws.column_dimensions["A"].width = 11
    for j, col in enumerate(px.columns, start=2):
        desc, src = meta[col]
        for r, val in ((1, col), (2, desc), (3, src)):
            c = ws.cell(r, j, val)
            c.font = F_HEAD if r == 1 else Font(name="Roboto Light", size=8, color="FFFFFF")
            c.fill = FILL_HEAD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 34
    for i, (dt, row) in enumerate(px.iterrows()):
        r = 4 + i
        c = ws.cell(r, 1, dt.to_pydatetime())
        c.number_format, c.font = DATEF, F_BOLD
        for j, col in enumerate(px.columns, start=2):
            val = row[col]
            c = ws.cell(r, j, None if pd.isna(val) else float(val))
            c.number_format = "#,##0.00"
            c.font = Font(name="Roboto Light", size=9, italic=True, color=GREY) if filled.at[dt, col] else F_BODY
    ws.freeze_panes = "B4"


BBG_ADJ_BR = ',"CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y"'   # retorno total, como os índices da B3
BBG_ADJ_MX = ',"CshAdjNormal=N","CshAdjAbnormal=N","CapChg=Y"'   # preço, como o MEXBOL


def bbg_series(bz: pd.DataFrame, mx: pd.DataFrame) -> list[tuple[str, str, str, str]]:
    """(código da aba Prices, ticker Bloomberg, tipo, descrição), na mesma ordem da versão estática."""
    out = []
    bbg_br = dict(zip(bz["ticker"], bz["bbg"]))
    for s in dict.fromkeys(bz["series"]):
        src = next((k for k, v in BDR.items() if v == s), None)
        tk = f"{s} BS Equity" if src else bbg_br[s]
        out.append((s, tk, "br", f"{s} BDR of {src} (BRL)" if src else f"{s} (BRL)"))
    info = {c: (n, b.replace(" (to confirm)", "")) for c, n, b, _ in CONTROL_INFO}
    for c in B3_INDICES:
        out.append((c, info[c][1], "idx", info[c][0]))
    for s, t, bt in zip(mx["series"], mx["ticker"], mx["security_ticker"]):
        out.append((s, bt, "mx", f"{t} (MXN)"))
    out.append(("MEXBOL", "MEXBOL Index", "idx", "S&P/BMV IPC (MXN)"))
    return out


def write_prices_bbg(wb, series: list[tuple[str, str, str, str]]):
    """Aba Prices da versão Bloomberg: uma linha por data-chave (corte, anúncio, dia seguinte, último) e uma coluna
    por série, com BDH de data única. São 4 pontos por série (~300 no total) e nenhuma fórmula de matriz.
    Último preço: com LastDate >= hoje, BDP (preço atual, intradiário com o mercado aberto)."""
    ws = wb.create_sheet("Prices")
    ws.sheet_view.showGridLines = False
    for r, lab in ((1, "Series"), (2, "Bloomberg ticker"), (3, "Description")):
        c = ws.cell(r, 1, lab)
        c.font, c.fill = F_HEAD, FILL_HEAD
    ws.column_dimensions["A"].width = 16
    for j, (code, tk, kind, desc) in enumerate(series, start=2):
        for r, val in ((1, code), (2, tk), (3, desc)):
            c = ws.cell(r, j, val)
            c.font = F_HEAD if r == 1 else Font(name="Roboto Light", size=8, color="FFFFFF")
            c.fill = FILL_HEAD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(2, j).fill = PatternFill("solid", fgColor="2E4A6B")  # ticker editável
        ws.column_dimensions[get_column_letter(j)].width = 12
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 34
    adj = {"br": BBG_ADJ_BR, "mx": BBG_ADJ_MX, "idx": ""}
    for i, nm in enumerate(["CutoffDate", "AnnDate", "DayAfterDate", "LastDate"]):
        r = 4 + i
        c = ws.cell(r, 1, f"={nm}")
        c.number_format, c.font = DATEF, F_BOLD
        for j, (code, tk, kind, desc) in enumerate(series, start=2):
            col = get_column_letter(j)
            hist = f'BDH({col}$2,"PX_LAST",$A{r},$A{r},"Days=A","Fill=P"{adj[kind]})'
            f = f'=IF($A{r}>=TODAY(),BDP({col}$2,"PX_LAST"),{hist})' if nm == "LastDate" else f"={hist}"
            c = ws.cell(r, j, f)
            c.number_format, c.font = "#,##0.00", F_BODY
    ws.cell(9, 1, "Bloomberg formulas (BDH at each key date; BDP for the latest price). Edit a ticker in row 2 if it "
                  "shows #N/A. Brazil: prices adjusted for dividends (total return, like the B3 indices); Mexico: "
                  "unadjusted (price return, like MEXBOL).").font = F_NOTE
    ws.freeze_panes = "B4"


def write_basis(wb, last: date | None, aum: dict, aum_mx: float, aum_mx_src: str, n_br: int, n_mx: int,
                bbg: bool = False):
    ws = wb.create_sheet("Basis")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 70
    ws["A1"] = "Basis: dates, parameters, control groups and sources"
    ws["A1"].font = F_TITLE
    ws["A2"] = "Yellow cells are inputs; everything on the MSCI Brazil and MSCI Mexico tabs recalculates from them."
    ws["A2"].font = F_SUB

    def head(r, labels):
        for j, t in enumerate(labels, start=1):
            c = ws.cell(r, j, t)
            c.font, c.fill, c.border = F_HEAD, FILL_HEAD, B_HEAD

    head(4, ["Key dates", "Value", "Name", "Note"])
    dates = [
        ("MSCI price cutoff", CUTOFF, "CutoffDate", "Price date MSCI used for the review; matches the pro forma "
         "local prices for Brazil (median error 0.0%) and Mexico (0.2%)"),
        ("Announcement", ANN, "AnnDate", "MSCI announced after the close; returns start at this close"),
        ("Day after", DAY_AFTER, "DayAfterDate", "First session after the announcement"),
        ("Last price", last, "LastDate",
         "=TODAY() takes the latest Bloomberg price (live while the market is open); type a date to freeze it"
         if bbg else "Change this to update all returns (needs a row on the Prices tab)"),
        ("Implementation (close)", IMPL, "ImplDate", f"Index changes take effect on {EFFECTIVE:%d-%b-%Y}"),
    ]
    for i, (lab, d, nm, note) in enumerate(dates):
        r = 5 + i
        ws.cell(r, 1, lab).font = F_BOLD
        c = ws.cell(r, 2, "=TODAY()" if d is None else datetime(d.year, d.month, d.day))
        c.number_format, c.font, c.fill = DATEF, F_BOLD, FILL_INPUT
        ws.cell(r, 3, nm).font = F_NOTE
        ws.cell(r, 4, note).font = F_BODY
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Basis!$B${r}")

    r0 = 11
    head(r0, ["Flow parameters", "Value", "Name", "Note"])
    params = [
        ("MSCI Brazil passive AUM (US$ mn)", aum["Brazil"][1], "AUM_BR",
         "Lipper, funds benchmarked to MSCI Brazil (same base as the published flows report of 11-Nov-2026)"),
        ("MSCI Mexico passive AUM (US$ mn)", aum_mx, "AUM_MX", aum_mx_src),
        ("MSCI EM Latin America passive AUM (US$ mn)", aum["Latam"][1], "AUM_LATAM",
         "Lipper, passive funds benchmarked to MSCI EM Latin America (flows report of 11-Nov-2026)"),
        ("MSCI EM passive AUM (US$ mn)", aum["EM"][1], "AUM_EM",
         "Lipper, passive funds benchmarked to the MSCI EM Standard index (flows report of 11-Nov-2026)"),
        ("MSCI ACWI passive AUM (US$ mn)", aum["ACWI"][1], "AUM_ACWI",
         "Lipper, passive funds benchmarked to MSCI ACWI (flows report of 11-Nov-2026)"),
        ("EM weight in the ACWI", em_in_acwi(), "EM_IN_ACWI",
         "A stock's change in weight in the ACWI = its change in the EM x this weight"),
        ("Active flows as a multiple of passive", 0.5, "ActiveRatio",
         "House convention: total = passive x 1.5 (calibrated on implementation-day volume, Feb/25 to Aug/26)"),
    ]
    for i, (lab, val, nm, note) in enumerate(params):
        r = r0 + 1 + i
        ws.cell(r, 1, lab).font = F_BOLD
        c = ws.cell(r, 2, val)
        c.number_format = "#,##0.0" if val > 5 else "0.00%" if nm == "EM_IN_ACWI" else "0.00"
        c.font, c.fill = F_BOLD, FILL_INPUT
        ws.cell(r, 3, nm).font = F_NOTE
        ws.cell(r, 4, note).font = F_BODY
        ws.cell(r, 4).alignment = Alignment(wrap_text=True, vertical="top")
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Basis!$B${r}")
    r = r0 + len(params)

    r1 = r + 2
    head(r1, ["Control groups", "B3 / Yahoo code", "Bloomberg ticker", "Used for"])
    for i, (code, name, bbg, used) in enumerate(CONTROL_INFO):
        rr = r1 + 1 + i
        ws.cell(rr, 1, name).font = F_BOLD
        ws.cell(rr, 2, code).font = F_BODY
        ws.cell(rr, 3, bbg).font = F_BODY
        ws.cell(rr, 4, used).font = F_BODY
        if i % 2:
            for j in range(1, 5):
                ws.cell(rr, j).fill = FILL_ZEBRA

    r2 = r1 + len(CONTROL_INFO) + 2
    head(r2, ["Method and sources", "", "", ""])
    notes = [
        ("Universe", f"MSCI Brazil: {n_br} names (the {n_br - 1} members plus JBS, deleted). MSCI Mexico: {n_mx} members "
                     "(no additions or deletions this cycle)."),
        ("Weights", "MSCI pro forma file of the November 2026 Index Review (prices of 20-Oct-2026). Mexico weights are "
                    "within MSCI Mexico (index market cap over the sum of the 21 members)."),
        ("Flows", "Each group of passive funds trades the change in the stock's weight in its own index: country "
                  "funds = change in weight in MSCI Brazil/Mexico x country AUM; regional & global = change in weight "
                  "in MSCI EM Latin America x LatAm AUM + change in weight in MSCI EM x (EM AUM + EM weight in ACWI x "
                  "ACWI AUM). Total = passive x (1 + active multiple). Impact = flow / 3M ADTV."),
        ("vs. the report", "The flows report of 11-Nov-2026 scaled the regional and global funds by the country's weight "
                           "(as if they reinvested in the country what they sell of a deleted name). Here they reinvest "
                           "across their whole index, as they do. Names with index events barely change (JBS, AXIA3, "
                           "TIMS3, GMEXICOB); names without events get smaller inflows (e.g., PETR4), and the country's "
                           "net flow follows its change in weight in the EM (Brazil 5.03% to 5.00%)."),
        ("JBS", "Net of the whole migration, as published: exit from the Standard indices and from EM IMI and ACWI ex "
                "USA IMI funds, minus purchases by MSCI USA Small Cap, USA IMI and World Small Cap funds "
                "(passive -302 mn; total -453 mn)."),
        ("ADTV", "3-month (63 sessions to 11-Nov-2026) average daily traded value in US$ mn: Economatica for B3, "
                 "Yahoo consolidated volume for US-listed names, LSEG/Yahoo for Mexico."),
        ("Returns", "Brazil: total return (Yahoo close adjusted for dividends) against B3 sector indices, which are "
                    "total return. US-listed names (NU, XP, AUGO, JBS) use their B3 BDRs (ROXO34, XPBR31, AURA33, "
                    "JBSS32), so stock and control are in BRL and close at the same time. Mexico: price return "
                    "(close) against MEXBOL, a price index."),
        ("Control groups", "Each Brazilian name goes to the B3 sector index that holds it in its theoretical portfolio "
                           "(16-Nov-2026). Names outside all of them take the index of their sector group. Sanitation "
                           "uses UTIL; oil & gas and telecom, with no B3 sector index, use the Ibovespa; JBS uses MBRF3. "
                           "The control index includes the stock itself (e.g., AXIA3 is ~14% of IEE), which mutes the "
                           "relative return. To change a control group, type another code from the Prices tab in the "
                           "'Control group' column."),
        ("This cycle", "The window since the 20-Oct cutoff includes the 26-Oct post-election sell-off (Ibovespa -13.9%, "
                       "circuit breaker; BRL -8%). Domestic sector indices fell 17% to 30% and exporters and US$ assets "
                       "held up, so relative returns since the cutoff mostly reflect the election, not the index "
                       "flows. The since-announcement and day-after windows are cleaner."),
        ("Updating", "Prices tab: append one row per trading day below the last one (oldest to newest), then change "
                     "LastDate above. Grey italics on the Prices tab = no trading that day (holiday), last close "
                     "carried forward. Or rerun build_charts_preview.py --last YYYY-MM-DD."),
        ("Sources", "MSCI; Lipper; Economatica; B3; Yahoo Finance; XP Research."),
    ]
    if bbg:
        notes = [(k, t) for k, t in notes if k not in ("Returns", "Updating", "Sources")] + [
            ("Returns", "Bloomberg PX_LAST at each key date. Brazil: adjusted for dividends (total return) against B3 "
                        "sector indices, which are total return. US-listed names (NU, XP, AUGO, JBS) use their B3 BDRs "
                        "(ROXO34, XPBR31, AURA33, JBSS32), so stock and control are in BRL and close at the same "
                        "time. Mexico: unadjusted (price return) against MEXBOL, a price index."),
            ("Updating", "Opens and refreshes with the Bloomberg Excel add-in (terminal logged in). LastDate = TODAY() "
                         "uses the current price; type a trading day to freeze it. About 300 data points per refresh "
                         "(4 dates x ~75 series). A ticker that shows #N/A can be fixed in row 2 of the Prices tab."),
            ("Sources", "MSCI; Lipper; Economatica; Bloomberg; XP Research."),
        ]
    for i, (k, t) in enumerate(notes):
        rr = r2 + 1 + i
        ws.cell(rr, 1, k).font = F_BOLD
        ws.cell(rr, 1).alignment = Alignment(vertical="top")
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        c = ws.cell(rr, 2, t)
        c.font, c.alignment = F_BODY, Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[rr].height = 12.5 * max(1, int(np.ceil(len(t) / 105)))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.defined_names["PxDates"] = DefinedName(
        "PxDates", attr_text=f"OFFSET(Prices!$A$4,0,0,COUNT(Prices!$A$4:$A${PX_LAST_ROW}),1)")


def recalc_with_excel(path: Path):
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(path))
        xl.CalculateFull()
        wb.Worksheets(1).Activate()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--last", help="último pregão (AAAA-MM-DD); padrão: último fechamento oficial da B3")
    ap.add_argument("--aum-mx", type=float, help="passivo que segue o MSCI Mexico (US$ mn)")
    ap.add_argument("--bbg", action="store_true",
                    help="gera Charts_Preview_BBG.xlsx, com preços por fórmulas da Bloomberg (BDH/BDP)")
    a = ap.parse_args()
    last = date.fromisoformat(a.last) if a.last else None

    bm = base_msci()
    bz = brazil_table(bm)
    mx = mexico_table(bm)
    aum = aum_country()
    if a.aum_mx:
        aum_mx, aum_mx_src = a.aum_mx, "Input"
    else:
        aum_mx, aum_mx_src = aum_mexico()

    if not a.bbg:
        px, filled, meta, last = prices(bz, mx, last)
        DATA.mkdir(exist_ok=True)
        px.to_csv(DATA / f"prices_{last:%Y%m%d}.csv")

    # ordena pelo impacto total (em módulo), que é o que interessa a quem opera no dia
    w_acwi = em_in_acwi()
    for df in (bz, mx):
        df["p_regional"] = (df["d_la"] * aum["Latam"][1] + df["d_em"] * (aum["EM"][1] + w_acwi * aum["ACWI"][1])) / 1e4             + df["p_other"]
    bz["_imp"] = ((bz["delta_bps"] / 1e4 * aum["Brazil"][1] + bz["p_regional"]) * 1.5 / bz["adtv"]).abs()
    bz = bz.sort_values("_imp", ascending=False)
    mx["_imp"] = ((mx["delta_bps"] / 1e4 * aum_mx + mx["p_regional"]) * 1.5 / mx["adtv"]).abs()
    mx = mx.sort_values("_imp", ascending=False)

    wb = Workbook()
    wb.remove(wb.active)
    notes_common = [
        f"Source: MSCI, Lipper, Economatica, {'Bloomberg' if a.bbg else 'B3, Yahoo Finance'}, XP Research. Flows are "
        "XP Research estimates; negative = outflow. Relative return in percentage points.",
    ]
    write_market_tab(
        wb, "MSCI Brazil", bz, "AUM_BR",
        "Announced 11-Nov-2026 | effective 1-Dec-2026, traded at the close of 30-Nov-2026 | "
        "No additions; JBS deleted (reclassified as a US company)",
        notes_common + ["NU, XP, AUGO and JBS: returns from the B3 BDRs (ROXO34, XPBR31, AURA33, JBSS32). "
                        "Control group codes on the Basis tab."])
    write_market_tab(
        wb, "MSCI Mexico", mx, "AUM_MX",
        "Announced 11-Nov-2026 | effective 1-Dec-2026, traded at the close of 30-Nov-2026 | "
        "No additions or deletions; FIF changes in GFINBURO, PINFRA and GMEXICOB",
        notes_common + ["Country-fund flows: passive funds benchmarked to MSCI Mexico (Lipper, plus the iShares "
                        "MSCI Mexico Capped UCITS ETF); see Basis. Control group: MEXBOL (S&P/BMV IPC)."])
    if a.bbg:
        series = bbg_series(bz, mx)
        write_prices_bbg(wb, series)
        write_basis(wb, None, aum, aum_mx, aum_mx_src, len(bz), len(mx), bbg=True)
        wb.active = 0
        wb.calculation.fullCalcOnLoad = True
        out = OUT.with_name("Charts_Preview_BBG.xlsx")
        wb.save(out)  # sem recálculo aqui: as fórmulas da Bloomberg só respondem no Excel com o add-in
        print(f"ok: {out} | {len(series)} séries Bloomberg | BR {len(bz)} | MX {len(mx)}")
        return
    write_prices(wb, px, filled, meta)
    write_basis(wb, last, aum, aum_mx, aum_mx_src, len(bz), len(mx))
    wb.active = 0
    wb.save(OUT)
    recalc_with_excel(OUT)
    print(f"ok: {OUT} | last={last} | BR {len(bz)} | MX {len(mx)} | AUM_MX={aum_mx:,.1f}")


if __name__ == "__main__":
    main()
