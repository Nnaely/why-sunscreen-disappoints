from pathlib import Path
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_dashboard import THEMES, TONE_MAP, is_complaint  # noqa: E402  (definisi tema yang sama dengan notebook 03)

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "Portofolio_Ulasan_Skincare_Sephora.xlsx"

# ---------- gaya (mengikuti Portofolio_Penjualan_Ecommerce_Indonesia.xlsx) ----------
NAVY, GREY_TXT, INK2 = "24364B", "63758A", "4A5468"
BLUE, CONTEXT, LIGHT, INPUT = "2A78D6", "C5CBD3", "F3F5F8", "FFF4D6"
TONES_HEX = ["F3D6BF", "D9A882", "B07A4F", "6B4428"]
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), color=k.pop("color", NAVY), **k)
FILL = lambda c: PatternFill("solid", fgColor=c)
THIN = Side(style="thin", color="D5DAE1")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PCT, PCT0, INT, DEC2 = "0.0%", "0%", "#,##0", "0.00"

SKIN_ID = {"combination": "Kombinasi", "dry": "Kering", "normal": "Normal", "oily": "Berminyak", "unknown": "Tidak diketahui"}
SKINS = ["Berminyak", "Kombinasi", "Kering", "Normal", "Tidak diketahui"]
TONES = ["Terang", "Medium", "Sawo matang", "Gelap", "Tidak diketahui"]
BANDS = ["< $25", "$25-50", "$50-100", "> $100"]


def title_block(ws, title, subtitle, width_cols=10):
    ws.sheet_view.showGridLines = False
    ws["A1"].value = title
    ws["A1"].font = F(size=14, bold=True)
    ws["A2"].value = subtitle
    ws["A2"].font = F(color=GREY_TXT)


def header(ws, row, col, labels):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = FILL(NAVY)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def section(ws, row, text, note=None):
    ws.cell(row=row, column=1, value=text).font = F(size=12, bold=True)
    if note:
        ws.cell(row=row + 1, column=1, value=note).font = F(color=GREY_TXT)


def add_table(ws, name, ref):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


def rich(size=900, color=INK2, bold=False):
    cp = CharacterProperties(latin=DFont(typeface="Arial"), sz=size, b=bold, solidFill=color)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def style_chart(ch, numfmt, horizontal=False, legend=False):
    ch.style = 2
    ch.legend = None if not legend else ch.legend
    if legend:
        ch.legend.position = "t"
        ch.legend.txPr = rich()
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
        ax.txPr = rich()
        ax.majorTickMark = "none"
        ax.minorTickMark = "none"
    ch.y_axis.numFmt = numfmt
    ch.y_axis.majorGridlines.spPr = None
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E6E9EE"))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="C5CBD3"))
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"   # urutan teratas = nilai terbesar
        ch.y_axis.crosses = "max"
    return ch


def labels(numfmt, size=850):
    d = DataLabelList()
    d.showVal = True
    for a in ("showSerName", "showCatName", "showLegendKey", "showPercent", "showLeaderLines"):
        setattr(d, a, False)
    d.numFmt = numfmt
    d.txPr = rich(size)
    return d


def color_series(s, hexc):
    s.graphicalProperties.solidFill = hexc
    s.graphicalProperties.line.noFill = True


def main():
    # ------------------------------------------------------------------ data
    r = pd.read_parquet(ROOT / "data/processed/reviews_clean.parquet",
                        columns=["rating", "is_recommended", "skin_type", "skin_tone", "year", "secondary_category",
                                 "price_band", "product_name", "brand_name", "review_title", "review_text"])
    r["skin"] = r["skin_type"].map(SKIN_ID)
    r["tone"] = r["skin_tone"].map(TONE_MAP).fillna("Tidak diketahui")
    r["price_band"] = r["price_band"].astype(str)
    r["neg"] = (r["rating"] <= 2).astype(int)
    r["rec_yes"] = (r["is_recommended"] == 1).astype(int)
    r["rec_n"] = r["is_recommended"].notna().astype(int)
    total_all = len(r)

    cat_counts = r["secondary_category"].value_counts()
    cats_used = cat_counts[cat_counts >= 5000].index.tolist()
    ru = r[r["secondary_category"].isin(cats_used)]
    cube = (ru.groupby(["secondary_category", "price_band", "skin", "tone", "year"], observed=True)
            .agg(n=("rating", "size"), stars=("rating", "sum"), neg=("neg", "sum"),
                 rec_yes=("rec_yes", "sum"), rec_n=("rec_n", "sum")).reset_index())
    cat_order = (ru.groupby("secondary_category")["neg"].mean().sort_values(ascending=False).index.tolist())

    sun = r[r["secondary_category"] == "Sunscreen"].copy()
    text = (sun["review_title"].fillna("") + ". " + sun["review_text"].fillna("")).str.lower()
    themes = list(THEMES)
    for t in themes:
        sun[t] = text.map(lambda x, p=THEMES[t]: is_complaint(x, p)).astype(int)
    sun["sent"] = np.select([sun["rating"] <= 2, sun["rating"] >= 4], ["Negatif", "Positif"], "Netral")
    sun2 = sun[sun["sent"] != "Netral"]
    scube = sun2.groupby(["sent", "skin", "tone", "year"]).agg(n=("rating", "size"), **{t: (t, "sum") for t in themes}).reset_index()
    theme_order = sun2[sun2["sent"] == "Negatif"][themes].mean().sort_values(ascending=False).index.tolist()

    prod = (sun.groupby(["brand_name", "product_name"])
            .agg(n=("rating", "size"), stars=("rating", "sum"), neg=("neg", "sum"),
                 rec_yes=("rec_yes", "sum"), rec_n=("rec_n", "sum")).query("n >= 500").reset_index())
    negp = sun2[sun2["sent"] == "Negatif"].groupby("product_name")[themes].mean()
    prod["issue"] = prod["product_name"].map(lambda p: negp.loc[p].idxmax() if p in negp.index else "-")
    prod["issue_pct"] = prod["product_name"].map(lambda p: negp.loc[p].max() if p in negp.index else None)
    prod = prod.sort_values("neg", key=lambda s: s / prod["n"], ascending=False)

    years = sorted(int(y) for y in r["year"].unique())

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    an = wb.create_sheet("Analisis")
    du = wb.create_sheet("Data_Ulasan")
    ds = wb.create_sheet("Data_Sunscreen")
    dp = wb.create_sheet("Produk_Sunscreen")
    tm = wb.create_sheet("Temuan")
    kq = wb.create_sheet("Kualitas")
    kd = wb.create_sheet("Kamus_Data")
    doc = wb.create_sheet("Dokumentasi")

    # ------------------------------------------------------------------ Data_Ulasan
    title_block(du, "Data ulasan agregat", "Satu baris = kombinasi kategori × kelompok harga × jenis kulit × warna kulit × tahun. Sumber: 1.092.952 ulasan dari 12 kategori skincare.")
    H0 = 4
    header(du, H0, 1, ["Kategori", "Kelompok harga", "Jenis kulit", "Warna kulit", "Tahun", "Jumlah ulasan",
                        "Total bintang", "Ulasan negatif (1–2★)", "Rekomendasi: ya", "Rekomendasi: dijawab"])
    for i, row in enumerate(cube.itertuples(index=False), start=H0 + 1):
        for j, v in enumerate(row, start=1):
            du.cell(row=i, column=j, value=(int(v) if isinstance(v, (np.integer,)) else v)).font = F()
    U_END = H0 + len(cube)
    add_table(du, "TblUlasan", f"A{H0}:J{U_END}")
    for c, w in zip("ABCDEFGHIJ", [24, 15, 16, 16, 8, 14, 14, 16, 15, 17]):
        du.column_dimensions[c].width = w
    du.freeze_panes = f"A{H0 + 1}"
    UR = lambda col: f"Data_Ulasan!${col}${H0 + 1}:${col}${U_END}"

    # ------------------------------------------------------------------ Data_Sunscreen
    title_block(ds, "Data keluhan sunscreen agregat", "Satu baris = sentimen × jenis kulit × warna kulit × tahun. Kolom tema = jumlah ulasan yang mengeluhkan tema itu (negasi seperti “no white cast” tidak dihitung).")
    header(ds, H0, 1, ["Sentimen", "Jenis kulit", "Warna kulit", "Tahun", "Jumlah ulasan"] + themes)
    for i, row in enumerate(scube.itertuples(index=False), start=H0 + 1):
        for j, v in enumerate(row, start=1):
            ds.cell(row=i, column=j, value=(int(v) if isinstance(v, (np.integer,)) else v)).font = F()
    S_END = H0 + len(scube)
    add_table(ds, "TblSunscreen", f"A{H0}:{L(5 + len(themes))}{S_END}")
    for j in range(1, 6 + len(themes)):
        ds.column_dimensions[L(j)].width = 14 if j > 5 else 15
    ds.freeze_panes = f"A{H0 + 1}"
    SR = lambda col: f"Data_Sunscreen!${col}${H0 + 1}:${col}${S_END}"
    THEME_COL = {t: L(6 + i) for i, t in enumerate(themes)}

    # ------------------------------------------------------------------ Dashboard: filter cells (dibuat dulu agar bisa dirujuk)
    FS, FT, FY0, FY1 = "Dashboard!$B$8", "Dashboard!$E$8", "Dashboard!$H$8", "Dashboard!$J$8"

    # ------------------------------------------------------------------ Analisis
    title_block(an, "Analisis ulasan skincare Sephora", "Semua angka dihitung dengan SUMIFS dari Data_Ulasan dan Data_Sunscreen, mengikuti filter di Dashboard.")
    an.column_dimensions["A"].width = 30
    for c in "BCDEFGHIJ":
        an.column_dimensions[c].width = 15

    row = 4
    section(an, row, "Parameter filter", "Diambil dari sel kuning di Dashboard. “*” berarti semua.")
    header(an, row + 2, 1, ["Parameter", "Kriteria"])
    params = [("Jenis kulit", f'=IF({FS}="Semua","*",{FS})'), ("Warna kulit", f'=IF({FT}="Semua","*",{FT})'),
              ("Tahun awal", f"=MIN({FY0},{FY1})"), ("Tahun akhir", f"=MAX({FY0},{FY1})")]
    for i, (k, f) in enumerate(params):
        an.cell(row=row + 3 + i, column=1, value=k).font = F()
        an.cell(row=row + 3 + i, column=2, value=f).font = F()
    P_SKIN, P_TONE, P_Y0, P_Y1 = "$B$7", "$B$8", "$B$9", "$B$10"
    crit = (f'{UR("C")},{P_SKIN},{UR("D")},{P_TONE},{UR("E")},">="&{P_Y0},{UR("E")},"<="&{P_Y1}')
    crit_s = lambda skin=True, tone=True: (
        (f',{SR("B")},{P_SKIN}' if skin else "") + (f',{SR("C")},{P_TONE}' if tone else "")
        + f',{SR("D")},">="&{P_Y0},{SR("D")},"<="&{P_Y1}')

    # --- Metrik utama
    row = 13
    section(an, row, "Metrik utama")
    header(an, row + 1, 1, ["Metrik", "Nilai", "Definisi"])
    an.column_dimensions["C"].width = 15
    km = [
        ("Ulasan dianalisis", f"=SUMIFS({UR('F')},{crit})", INT, "Jumlah ulasan pada filter aktif"),
        ("Rating rata-rata", f"=IFERROR(SUMIFS({UR('G')},{crit})/B15,\"n.a.\")", DEC2, "Total bintang ÷ jumlah ulasan"),
        ("Ulasan negatif", f"=IFERROR(SUMIFS({UR('H')},{crit})/B15,\"n.a.\")", PCT, "Rating 1–2 bintang ÷ jumlah ulasan"),
        ("Merekomendasikan", f"=IFERROR(SUMIFS({UR('I')},{crit})/SUMIFS({UR('J')},{crit}),\"n.a.\")", PCT, "Jawaban ‘ya’ ÷ pengulas yang menjawab"),
        ("Ulasan sunscreen", f"=SUMIFS({UR('F')},{UR('A')},\"Sunscreen\",{crit})", INT, "Jumlah ulasan kategori Sunscreen"),
        ("Ulasan negatif sunscreen", f"=IFERROR(SUMIFS({UR('H')},{UR('A')},\"Sunscreen\",{crit})/B19,\"n.a.\")", PCT, "Rating 1–2 bintang ÷ ulasan sunscreen"),
        ("Sunscreen dibanding rata-rata", "=IFERROR(B20/B17,\"n.a.\")", '0.0"×"', "Ulasan negatif sunscreen ÷ ulasan negatif semua kategori"),
    ]
    for i, (k, f, fmt, d) in enumerate(km):
        rr = row + 2 + i
        an.cell(row=rr, column=1, value=k).font = F()
        c = an.cell(row=rr, column=2, value=f); c.font = F(); c.number_format = fmt
        an.cell(row=rr, column=3, value=d).font = F(color=GREY_TXT)
    assert row + 2 == 15
    for j, lab in ((4, "Semua data"), (5, "Selisih / bagian")):
        c = an.cell(row=14, column=j, value=lab); c.font = F(bold=True, color="FFFFFF"); c.fill = FILL(NAVY)
    alld = [
        (f"=SUM({UR('F')})", INT, "=B15/D15", PCT),
        (f"=SUM({UR('G')})/D15", DEC2, "=IFERROR(B16-D16,0)", '+0.00;-0.00;0.00'),
        (f"=SUM({UR('H')})/D15", PCT, "=IFERROR((B17-D17)*100,0)", '+0.0" poin";-0.0" poin";0.0" poin"'),
        (f"=SUM({UR('I')})/SUM({UR('J')})", PCT, "=IFERROR((B18-D18)*100,0)", '+0.0" poin";-0.0" poin";0.0" poin"'),
        (f'=SUMIFS({UR("F")},{UR("A")},"Sunscreen")', INT, "=IFERROR(B19/D19,0)", PCT),
        (f'=SUMIFS({UR("H")},{UR("A")},"Sunscreen")/D19', PCT, "=IFERROR((B20-D20)*100,0)", '+0.0" poin";-0.0" poin";0.0" poin"'),
        ("=D20/D17", '0.0"×"', None, None),
    ]
    for i, (fa, fmta, fb, fmtb) in enumerate(alld):
        c = an.cell(row=15 + i, column=4, value=fa); c.font = F(); c.number_format = fmta
        if fb:
            c = an.cell(row=15 + i, column=5, value=fb); c.font = F(); c.number_format = fmtb

    # --- Per kategori
    row = 24
    section(an, row, "Ulasan negatif per kategori", "Tabel kiri dihitung per kategori; tabel kanan mengurutkannya untuk grafik (Sunscreen dipisah agar bisa disorot).")
    header(an, row + 2, 1, ["Kategori", "Jumlah ulasan", "Ulasan negatif", "% negatif", "Rating rata-rata", "Urutan bantu"])
    c0 = row + 3
    for i, cat in enumerate(cat_order):
        rr = c0 + i
        an.cell(row=rr, column=1, value=cat).font = F()
        an.cell(row=rr, column=2, value=f"=SUMIFS({UR('F')},{UR('A')},\"=\"&$A{rr},{crit})").number_format = INT
        an.cell(row=rr, column=3, value=f"=SUMIFS({UR('H')},{UR('A')},\"=\"&$A{rr},{crit})").number_format = INT
        an.cell(row=rr, column=4, value=f"=IF(B{rr}=0,0,C{rr}/B{rr})").number_format = PCT
        an.cell(row=rr, column=5, value=f"=IF(B{rr}=0,\"n.a.\",SUMIFS({UR('G')},{UR('A')},\"=\"&$A{rr},{crit})/B{rr})").number_format = DEC2
        an.cell(row=rr, column=6, value=f"=D{rr}+ROW()/1E9").number_format = "0.000000"
        for j in range(2, 7):
            an.cell(row=rr, column=j).font = F()
    c1 = c0 + len(cat_order) - 1
    header(an, row + 2, 8, ["Kategori (urut)", "Sunscreen", "Kategori lain", "Jumlah ulasan", "% negatif"])
    an.column_dimensions["H"].width = 24
    for k in range(len(cat_order)):
        rr = c0 + k
        an.cell(row=rr, column=8, value=f"=INDEX($A${c0}:$A${c1},MATCH(LARGE($F${c0}:$F${c1},{k + 1}),$F${c0}:$F${c1},0))").font = F()
        v = f"INDEX($D${c0}:$D${c1},MATCH($H{rr},$A${c0}:$A${c1},0))"
        an.cell(row=rr, column=9, value=f'=IF($H{rr}="Sunscreen",{v},0)').number_format = '0.0%;-0.0%;;'
        an.cell(row=rr, column=10, value=f'=IF($H{rr}<>"Sunscreen",{v},0)').number_format = '0.0%;-0.0%;;'
        an.cell(row=rr, column=9).font = F(); an.cell(row=rr, column=10).font = F()
        c = an.cell(row=rr, column=11, value=f"=INDEX($B${c0}:$B${c1},MATCH($H{rr},$A${c0}:$A${c1},0))"); c.font = F(); c.number_format = INT
        c = an.cell(row=rr, column=12, value=f"=INDEX($D${c0}:$D${c1},MATCH($H{rr},$A${c0}:$A${c1},0))"); c.font = F(); c.number_format = PCT
    CAT = (c0, c1)

    # --- Per kelompok harga
    row = c1 + 3
    section(an, row, "Rekomendasi per kelompok harga")
    header(an, row + 1, 1, ["Kelompok harga", "Jumlah ulasan", "Rekomendasi: ya", "Rekomendasi: dijawab", "% merekomendasikan", "% negatif"])
    p0 = row + 2
    for i, b in enumerate(BANDS):
        rr = p0 + i
        an.cell(row=rr, column=1, value=b).font = F()
        an.cell(row=rr, column=2, value=f"=SUMIFS({UR('F')},{UR('B')},\"=\"&$A{rr},{crit})").number_format = INT
        an.cell(row=rr, column=3, value=f"=SUMIFS({UR('I')},{UR('B')},\"=\"&$A{rr},{crit})").number_format = INT
        an.cell(row=rr, column=4, value=f"=SUMIFS({UR('J')},{UR('B')},\"=\"&$A{rr},{crit})").number_format = INT
        an.cell(row=rr, column=5, value=f"=IF(D{rr}=0,0,C{rr}/D{rr})").number_format = PCT
        an.cell(row=rr, column=6, value=f"=IF(B{rr}=0,0,SUMIFS({UR('H')},{UR('B')},\"=\"&$A{rr},{crit})/B{rr})").number_format = PCT
        for j in range(2, 7):
            an.cell(row=rr, column=j).font = F()
    PRICE = (p0, p0 + len(BANDS) - 1)

    # --- Per tahun
    row = PRICE[1] + 3
    section(an, row, "Tren per tahun", "Dipakai untuk grafik tahunan dan sparkline. Tahun di luar filter bernilai 0 / #N/A. Data 2023 hanya Januari–Maret.")
    header(an, row + 2, 1, ["Tahun", "Jumlah ulasan", "Ulasan negatif", "% negatif", "Ulasan lainnya", "Rating rata-rata",
                             "% merekomendasikan", "Ulasan sunscreen", "% negatif sunscreen"])
    crit_noyear = f'{UR("C")},{P_SKIN},{UR("D")},{P_TONE}'
    y0r = row + 3
    for i, y in enumerate(years):
        rr = y0r + i
        an.cell(row=rr, column=1, value=str(y)).font = F()
        inr = f"AND({y}>={P_Y0},{y}<={P_Y1})"
        cy = f"{UR('E')},{y},{crit_noyear}"
        vals = [
            (f"=IF({inr},SUMIFS({UR('F')},{cy}),0)", INT),
            (f"=IF({inr},SUMIFS({UR('H')},{cy}),0)", INT),
            (f"=IF(B{rr}<100,NA(),C{rr}/B{rr})", PCT),
            (f"=B{rr}-C{rr}", INT),
            (f"=IF(B{rr}<100,NA(),SUMIFS({UR('G')},{cy})/B{rr})", DEC2),
            (f"=IF(B{rr}<100,NA(),IFERROR(SUMIFS({UR('I')},{cy})/SUMIFS({UR('J')},{cy}),NA()))", PCT),
            (f"=IF({inr},SUMIFS({UR('F')},{UR('A')},\"Sunscreen\",{cy}),0)", INT),
            (f"=IF(H{rr}<50,NA(),SUMIFS({UR('H')},{UR('A')},\"Sunscreen\",{cy})/H{rr})", PCT),
        ]
        for j, (f, fmt) in enumerate(vals, start=2):
            c = an.cell(row=rr, column=j, value=f); c.font = F(); c.number_format = fmt
    YEAR = (y0r, y0r + len(years) - 1)
    an.cell(row=YEAR[1] + 1, column=1, value="Rasio hanya ditampilkan bila tahun itu punya ≥ 100 ulasan (sunscreen ≥ 50). #N/A sengaja dipakai agar titik itu tidak digambar di grafik.").font = F(color=GREY_TXT)

    # --- Keluhan sunscreen
    row = YEAR[1] + 4
    section(an, row, "Keluhan di ulasan sunscreen", "Lift = % di ulasan negatif ÷ % di ulasan positif. Lift > 1 berarti tema itu khas ulasan negatif.")
    header(an, row + 2, 1, ["Tema keluhan", "Ulasan negatif menyebut", "Ulasan positif menyebut", "% di ulasan negatif", "% di ulasan positif", "Lift"])
    nneg = f'SUMIFS({SR("E")},{SR("A")},"Negatif"{crit_s()})'
    npos = f'SUMIFS({SR("E")},{SR("A")},"Positif"{crit_s()})'
    t0 = row + 3
    for i, t in enumerate(theme_order):
        rr = t0 + i
        col = THEME_COL[t]
        an.cell(row=rr, column=1, value=t).font = F()
        an.cell(row=rr, column=2, value=f'=SUMIFS({SR(col)},{SR("A")},"Negatif"{crit_s()})').number_format = INT
        an.cell(row=rr, column=3, value=f'=SUMIFS({SR(col)},{SR("A")},"Positif"{crit_s()})').number_format = INT
        an.cell(row=rr, column=4, value=f"=IFERROR(B{rr}/$B${t0 + len(themes) + 1},0)").number_format = PCT
        an.cell(row=rr, column=5, value=f"=IFERROR(C{rr}/$C${t0 + len(themes) + 1},0)").number_format = PCT
        an.cell(row=rr, column=6, value=f'=IFERROR(D{rr}/E{rr},"n.a.")').number_format = '0.0"×"'
        for j in range(2, 7):
            an.cell(row=rr, column=j).font = F()
    t1 = t0 + len(themes) - 1
    an.cell(row=t1 + 2, column=1, value="Total ulasan (penyebut)").font = F(bold=True)
    an.cell(row=t1 + 2, column=2, value=f"={nneg}").number_format = INT
    an.cell(row=t1 + 2, column=3, value=f"={npos}").number_format = INT
    THEME = (t0, t1)

    # --- White cast per warna kulit (abaikan filter warna kulit)
    row = t1 + 5
    section(an, row, "Keluhan white cast menurut warna kulit", "Mengikuti filter jenis kulit dan tahun, tetapi tidak filter warna kulit (agar keempat kelompok bisa dibandingkan).")
    header(an, row + 2, 1, ["Warna kulit", "Ulasan negatif", "Menyebut white cast", "% white cast"])
    w0 = row + 3
    for i, tn in enumerate(TONES[:4]):
        rr = w0 + i
        an.cell(row=rr, column=1, value=tn).font = F()
        an.cell(row=rr, column=2, value=f'=SUMIFS({SR("E")},{SR("A")},"Negatif",{SR("C")},$A{rr}{crit_s(tone=False)})').number_format = INT
        an.cell(row=rr, column=3, value=f'=SUMIFS({SR(THEME_COL["White cast"])},{SR("A")},"Negatif",{SR("C")},$A{rr}{crit_s(tone=False)})').number_format = INT
        an.cell(row=rr, column=4, value=f"=IF(B{rr}<30,0,C{rr}/B{rr})").number_format = PCT
        for j in range(2, 5):
            an.cell(row=rr, column=j).font = F()
    WC = (w0, w0 + 3)
    an.cell(row=WC[1] + 1, column=1, value="Kelompok dengan < 30 ulasan negatif ditampilkan 0 karena terlalu sedikit.").font = F(color=GREY_TXT)

    # --- Keluhan x jenis kulit (abaikan filter jenis kulit)
    row = WC[1] + 4
    section(an, row, "Keluhan sunscreen per jenis kulit", "% ulasan negatif yang menyebut tema. Mengikuti filter warna kulit dan tahun, tetapi tidak filter jenis kulit.")
    header(an, row + 2, 1, ["Tema keluhan"] + SKINS[:4])
    hx0 = row + 3
    an.cell(row=hx0, column=1, value="Ulasan negatif (n)").font = F(bold=True)
    for j, sk in enumerate(SKINS[:4], start=2):
        an.cell(row=hx0, column=j, value=f'=SUMIFS({SR("E")},{SR("A")},"Negatif",{SR("B")},"{sk}"{crit_s(skin=False)})').number_format = INT
    for i, t in enumerate(theme_order[:7]):
        rr = hx0 + 1 + i
        an.cell(row=rr, column=1, value=t).font = F()
        for j, sk in enumerate(SKINS[:4], start=2):
            cnt = f'SUMIFS({SR(THEME_COL[t])},{SR("A")},"Negatif",{SR("B")},"{sk}"{crit_s(skin=False)})'
            c = an.cell(row=rr, column=j, value=f"=IF({L(j)}${hx0}<30,\"n.a.\",{cnt}/{L(j)}${hx0})")
            c.number_format = PCT0; c.font = F()
    HEAT = (hx0, hx0 + 7)

    row = HEAT[1] + 3
    section(an, row, "Pengulas menurut jenis kulit", "Mengikuti filter warna kulit dan tahun, tetapi tidak filter jenis kulit.")
    header(an, row + 2, 1, ["Jenis kulit", "Jumlah ulasan", "% dari total"])
    k0 = row + 3
    crit_noskin = f'{UR("D")},{P_TONE},{UR("E")},">="&{P_Y0},{UR("E")},"<="&{P_Y1}'
    for i, sk in enumerate(SKINS):
        rr = k0 + i
        an.cell(row=rr, column=1, value=sk).font = F()
        c = an.cell(row=rr, column=2, value=f'=SUMIFS({UR("F")},{UR("C")},$A{rr},{crit_noskin})'); c.font = F(); c.number_format = INT
        c = an.cell(row=rr, column=3, value=f"=IFERROR(B{rr}/$B${k0 + len(SKINS)},0)"); c.font = F(); c.number_format = PCT0
    an.cell(row=k0 + len(SKINS), column=1, value="Total").font = F(bold=True)
    c = an.cell(row=k0 + len(SKINS), column=2, value=f"=SUM(B{k0}:B{k0 + len(SKINS) - 1})"); c.font = F(bold=True); c.number_format = INT
    SKC = (k0, k0 + len(SKINS) - 1)
    an.freeze_panes = "A4"

    # ------------------------------------------------------------------ Produk_Sunscreen
    title_block(dp, "Produk sunscreen populer (≥ 500 ulasan)", "Seluruh periode, tidak terpengaruh filter. Persentase dihitung dengan rumus dari kolom jumlah.")
    header(dp, H0, 1, ["Brand", "Produk", "Jumlah ulasan", "Total bintang", "Ulasan negatif", "Rekomendasi: ya", "Rekomendasi: dijawab",
                        "Rating rata-rata", "% negatif", "% merekomendasikan", "Keluhan utama di ulasan negatif", "% keluhan utama"])
    for i, p in enumerate(prod.itertuples(index=False), start=H0 + 1):
        vals = [p.brand_name, p.product_name, int(p.n), int(p.stars), int(p.neg), int(p.rec_yes), int(p.rec_n)]
        for j, v in enumerate(vals, start=1):
            dp.cell(row=i, column=j, value=v).font = F()
        dp.cell(row=i, column=8, value=f"=D{i}/C{i}").number_format = DEC2
        dp.cell(row=i, column=9, value=f"=E{i}/C{i}").number_format = PCT
        dp.cell(row=i, column=10, value=f"=IF(G{i}=0,\"n.a.\",F{i}/G{i})").number_format = PCT
        dp.cell(row=i, column=11, value=p.issue).font = F()
        c = dp.cell(row=i, column=12, value=float(round(p.issue_pct, 4)) if p.issue_pct is not None else None)
        c.number_format = PCT0
        for j in (3, 4, 5, 6, 7):
            dp.cell(row=i, column=j).number_format = INT
        for j in range(8, 13):
            dp.cell(row=i, column=j).font = F()
    P_END = H0 + len(prod)
    add_table(dp, "TblProduk", f"A{H0}:L{P_END}")
    for c, w in zip("ABCDEFGHIJKL", [17, 58, 13, 13, 13, 14, 15, 13, 11, 15, 30, 13]):
        dp.column_dimensions[c].width = w
    dp.freeze_panes = f"C{H0 + 1}"
    dp.cell(row=P_END + 2, column=1, value="Keluhan utama dihitung di Python dengan metode tema yang sama (notebook 03), karena butuh teks ulasan.").font = F(color=GREY_TXT)

    # ------------------------------------------------------------------ Dashboard
    ws = dash
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    TEAL, TEAL_LIGHT, ORANGE, SLATE, BANNER, PAGE = "2A9D8F", "CDEBE6", "E76F51", "37474F", "263238", "EEF1F4"
    LAST_COL, LAST_ROW = 18, 112
    ws.column_dimensions["A"].width = 2
    for j in range(2, LAST_COL):
        ws.column_dimensions[L(j)].width = 10.5
    ws.column_dimensions[L(LAST_COL)].width = 2
    for rr in range(1, LAST_ROW + 1):
        for cc in range(1, LAST_COL + 1):
            ws.cell(row=rr, column=cc).fill = FILL(PAGE)

    def fill(rng, color):
        for row_ in ws[rng]:
            for c in row_:
                c.fill = FILL(color)

    def card(rng):
        fill(rng, "FFFFFF")
        cells = ws[rng]
        edge = Side(style="thin", color="D9DEE4")
        for i, row_ in enumerate(cells):
            for j, c in enumerate(row_):
                c.border = Border(left=edge if j == 0 else None, right=edge if j == len(row_) - 1 else None,
                                  top=edge if i == 0 else None, bottom=edge if i == len(cells) - 1 else None)

    def put(addr, value, merge_to=None, **font):
        if merge_to:
            ws.merge_cells(f"{addr}:{merge_to}")
        c = ws[addr]
        c.value = value
        align = font.pop("align", None)
        fmt = font.pop("fmt", None)
        c.font = F(**font)
        if align:
            c.alignment = align
        if fmt:
            c.number_format = fmt
        return c

    def section_bar(row, text):
        fill(f"B{row}:Q{row}", SLATE)
        ws.row_dimensions[row].height = 24
        c = put(f"B{row}", text, size=11, bold=True, color="FFFFFF", align=Alignment(vertical="center", indent=1))
        c.border = Border(left=Side(style="thick", color=TEAL))

    def card_title(addr, text):
        put(addr, text, size=9, bold=True, color="546E7A", align=Alignment(vertical="center", indent=1))

    def plain_chart(ch):
        ch.roundedCorners = False
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.line import LineProperties
        ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
        ch.graphical_properties.noFill = True
        ch.plot_area.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
        ch.plot_area.graphicalProperties.noFill = True
        return ch

    def sparkline(col, anchor, fmt):
        ch = LineChart()
        ch.add_data(Reference(an, min_col=col, min_row=YEAR[0] - 1, max_row=YEAR[1]), titles_from_data=True)
        ch.set_categories(Reference(an, min_col=1, min_row=YEAR[0], max_row=YEAR[1]))
        s = ch.series[0]
        s.graphicalProperties.line.solidFill = TEAL
        s.graphicalProperties.line.width = 22000
        s.marker.symbol = "none"
        s.smooth = False
        ch.legend = None
        ch.x_axis.delete = True
        ch.y_axis.delete = True
        ch.y_axis.majorGridlines = None
        ch.y_axis.numFmt = fmt
        ch.display_blanks = "gap"
        plain_chart(ch)
        ch.width, ch.height = 6.0, 2.3
        ws.add_chart(ch, anchor)

    # --- banner
    fill("B2:Q4", BANNER)
    put("B2", "LAPORAN ULASAN SKINCARE SEPHORA", size=16, bold=True, color="FFFFFF", align=Alignment(vertical="bottom", indent=1))
    put("B3", "1,09 juta ulasan, Agustus 2008 – Maret 2023. Fokus: kenapa pembeli sunscreen kecewa, sebagai masukan produk untuk brand lokal.",
        color="B0BEC5", align=Alignment(vertical="top", indent=1))
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 20

    # --- filter
    card("B6:Q8")
    put("B6", "Filter: pilih di sel kuning. Semua angka, grafik, dan tabel di bawah dihitung ulang.", size=9, color=GREY_TXT,
        align=Alignment(vertical="center", indent=1))
    for addr, lab in [("B7", "Jenis kulit"), ("E7", "Warna kulit"), ("H7", "Tahun dari"), ("J7", "Tahun sampai")]:
        put(addr, lab, size=9, bold=True, color=INK2, align=Alignment(indent=1))
    for rng, val, opts in [("B8:C8", "Semua", ["Semua"] + SKINS), ("E8:F8", "Semua", ["Semua"] + TONES),
                           ("H8:H8", years[0], [str(y) for y in years]), ("J8:J8", years[-1], [str(y) for y in years])]:
        a, b = rng.split(":")
        if a != b:
            ws.merge_cells(rng)
        c = ws[a]
        c.value = val
        c.font = F(bold=True)
        c.fill = FILL(INPUT)
        c.border = BOX
        c.alignment = Alignment(horizontal="left", indent=1)
        dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=False,
                            showErrorMessage=True, errorTitle="Pilihan tidak valid", error="Pilih dari daftar.")
        ws.add_data_validation(dv)
        dv.add(a)
    put("L8", (f'=IF(B8="Semua","Semua jenis kulit","Kulit "&LOWER(B8))&", "&'
               f'IF(E8="Semua","semua warna kulit","warna kulit "&LOWER(E8))&", "&Analisis!B9&"–"&Analisis!B10'),
        merge_to="Q8", size=9, color=GREY_TXT, align=Alignment(horizontal="right", indent=1))

    # --- KPI card: label, angka besar, sparkline tren per tahun, pembanding
    def kpi_card(c1, c2, r0, label, value, fmt, spark_col, spark_fmt, left_lab, left_val, left_fmt,
                 right_lab, right_val, right_fmt, better_high=None, h=10):
        card(f"{c1}{r0}:{c2}{r0 + h}")
        card_title(f"{c1}{r0}", label)
        put(f"{c1}{r0 + 1}", value, merge_to=f"{c2}{r0 + 2}", size=24, bold=True, fmt=fmt,
            align=Alignment(horizontal="center", vertical="center"))
        sparkline(spark_col, f"{c1}{r0 + 3}", spark_fmt)
        put(f"{c1}{r0 + h - 2}", left_lab, size=8, color=GREY_TXT, align=Alignment(indent=1))
        put(f"{c2}{r0 + h - 2}", right_lab, size=8, color=GREY_TXT, align=Alignment(horizontal="right"))
        put(f"{c1}{r0 + h - 1}", left_val, size=9, bold=True, color=INK2, fmt=left_fmt, align=Alignment(horizontal="left", indent=1))
        rc = put(f"{c2}{r0 + h - 1}", right_val, size=9, bold=True, color=INK2, fmt=right_fmt, align=Alignment(horizontal="right"))
        if better_high is not None:
            good, bad = Font(name="Arial", size=9, bold=True, color="2E7D32"), Font(name="Arial", size=9, bold=True, color="C62828")
            ws.conditional_formatting.add(rc.coordinate, CellIsRule(operator="greaterThan", formula=["0.0001"], font=good if better_high else bad))
            ws.conditional_formatting.add(rc.coordinate, CellIsRule(operator="lessThan", formula=["-0.0001"], font=bad if better_high else good))

    PTS = '+0.0" poin";-0.0" poin";0.0" poin"'
    YR = "Tren per tahun"

    section_bar(10, "RINGKASAN SEMUA SKINCARE")
    kpi_card("B", "D", 12, "ULASAN DIANALISIS", "=Analisis!B15", INT, 2, INT,
             "Semua data", "=Analisis!D15", INT, "Bagian dari total", "=Analisis!E15", PCT)

    # stacked: ulasan per tahun, negatif vs lainnya
    card("E12:Q22")
    card_title("E12", "ULASAN PER TAHUN: NEGATIF (RATING 1–2) DAN LAINNYA")
    put("N12", "2023 hanya Januari–Maret", merge_to="Q12", size=8, color=GREY_TXT, align=Alignment(horizontal="right", indent=1))
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.gapWidth = 45
    ch.add_data(Reference(an, min_col=5, min_row=YEAR[0] - 1, max_row=YEAR[1]), titles_from_data=True)
    ch.add_data(Reference(an, min_col=3, min_row=YEAR[0] - 1, max_row=YEAR[1]), titles_from_data=True)
    ch.set_categories(Reference(an, min_col=1, min_row=YEAR[0], max_row=YEAR[1]))
    color_series(ch.series[0], TEAL); color_series(ch.series[1], ORANGE)
    style_chart(ch, "#,##0", legend=True)
    plain_chart(ch)
    ch.width, ch.height = 25.0, 5.9
    ws.add_chart(ch, "E13")

    kpi_card("B", "D", 24, "RATING RATA-RATA", "=Analisis!B16", DEC2, 6, DEC2,
             "Semua data", "=Analisis!D16", DEC2, "Selisih", "=Analisis!E16", '+0.00;-0.00;0.00', better_high=True)
    kpi_card("E", "G", 24, "ULASAN NEGATIF", "=Analisis!B17", PCT, 4, PCT0,
             "Semua data", "=Analisis!D17", PCT, "Selisih", "=Analisis!E17", PTS, better_high=False)
    kpi_card("H", "J", 24, "MEREKOMENDASIKAN", "=Analisis!B18", PCT, 7, PCT0,
             "Semua data", "=Analisis!D18", PCT, "Selisih", "=Analisis!E18", PTS, better_high=True)

    # pie: komposisi jenis kulit + tabel legenda
    card("K24:Q34")
    card_title("K24", "PENGULAS MENURUT JENIS KULIT")
    pie = PieChart()
    pie.add_data(Reference(an, min_col=2, min_row=SKC[0] - 1, max_row=SKC[1]), titles_from_data=True)
    pie.set_categories(Reference(an, min_col=1, min_row=SKC[0], max_row=SKC[1]))
    for i, hx in enumerate([TEAL, ORANGE, "2A78D6", "E9C46A", "B0B7C3"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hx
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        pie.series[0].dPt.append(pt)
    pl = DataLabelList(); pl.showPercent = True
    for a in ("showVal", "showSerName", "showCatName", "showLegendKey", "showLeaderLines"):
        setattr(pl, a, False)
    pl.numFmt = "0%"; pl.txPr = rich(800, "FFFFFF", True)
    pie.series[0].dLbls = pl
    pie.legend = None
    plain_chart(pie)
    pie.width, pie.height = 6.0, 5.2
    ws.add_chart(pie, "K25")
    put("N26", "Jenis kulit", merge_to="O26", size=8, bold=True, color=INK2)
    put("P26", "Ulasan", size=8, bold=True, color=INK2, align=Alignment(horizontal="right"))
    put("Q26", "%", size=8, bold=True, color=INK2, align=Alignment(horizontal="right"))
    for i in range(5):
        rr = 27 + i
        put(f"N{rr}", f"=Analisis!A{SKC[0] + i}", merge_to=f"O{rr}", size=9)
        put(f"P{rr}", f"=Analisis!B{SKC[0] + i}", size=9, fmt=INT, align=Alignment(horizontal="right"))
        put(f"Q{rr}", f"=Analisis!C{SKC[0] + i}", size=9, fmt=PCT0, align=Alignment(horizontal="right"))
        ws[f"N{rr}"].border = Border(left=Side(style="thick", color=[TEAL, ORANGE, "2A78D6", "E9C46A", "B0B7C3"][i]))
    put("N32", "Total", merge_to="O32", size=9, bold=True)
    put("P32", f"=Analisis!B{SKC[1] + 1}", size=9, bold=True, fmt=INT, align=Alignment(horizontal="right"))
    for col in "NOPQ":
        ws[f"{col}32"].border = Border(top=Side(style="thin", color="90A4AE"))
    put("N33", "Tidak memakai filter jenis kulit.", merge_to="Q33", size=8, color=GREY_TXT)

    # --- Sunscreen
    section_bar(37, "SUNSCREEN: APA YANG DIKELUHKAN PEMBELI")
    kpi_card("B", "D", 39, "ULASAN NEGATIF SUNSCREEN", "=Analisis!B20", PCT, 9, PCT0,
             "Rata-rata skincare", "=Analisis!B17", PCT, "Selisih", "=(Analisis!B20-Analisis!B17)*100", PTS, better_high=False, h=13)

    card("E39:J52")
    card_title("E39", "KELUHAN WHITE CAST MENURUT WARNA KULIT")
    put("E40", "% ulasan negatif sunscreen yang menyebut white cast", merge_to="J40", size=8, color=GREY_TXT, align=Alignment(indent=1))
    ch = BarChart(); ch.type = "col"; ch.gapWidth = 50
    ch.add_data(Reference(an, min_col=4, min_row=WC[0] - 1, max_row=WC[1]), titles_from_data=True)
    ch.set_categories(Reference(an, min_col=1, min_row=WC[0], max_row=WC[1]))
    s = ch.series[0]
    s.graphicalProperties.line.solidFill = "8A93A3"
    for i, hx in enumerate(TONES_HEX):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hx
        pt.graphicalProperties.line.solidFill = "8A93A3"
        s.dPt.append(pt)
    s.dLbls = labels("0%", 1000)
    style_chart(ch, "0%")
    ch.y_axis.scaling.min = 0
    ch.y_axis.majorGridlines = None
    ch.y_axis.delete = True
    plain_chart(ch)
    ch.width, ch.height = 12.4, 5.8
    ws.add_chart(ch, "E41")
    put("E52", "Tidak memakai filter warna kulit. Kelompok < 30 ulasan negatif ditampilkan 0.", merge_to="J52", size=8, color=GREY_TXT, align=Alignment(indent=1))

    # tabel keluhan dengan data bar
    card("K39:Q52")
    card_title("K39", "KELUHAN DI ULASAN SUNSCREEN")
    for addr, t, al in [("K41", "Tema keluhan", "left"), ("N41", "% negatif", "left"), ("P41", "% positif", "right"), ("Q41", "Lift", "right")]:
        put(addr, t, size=8, bold=True, color=INK2, align=Alignment(horizontal=al, indent=1 if al == "left" else 0))
    ws.merge_cells("K41:M41"); ws.merge_cells("N41:O41")
    for col in "KLMNOPQ":
        ws[f"{col}41"].border = Border(bottom=Side(style="thin", color="90A4AE"))
    for i in range(len(themes)):
        rr = 42 + i
        put(f"K{rr}", f"=Analisis!A{THEME[0] + i}", merge_to=f"M{rr}", size=9, align=Alignment(indent=1))
        put(f"N{rr}", f"=Analisis!D{THEME[0] + i}", merge_to=f"O{rr}", size=9, bold=True, fmt=PCT)
        put(f"P{rr}", f"=Analisis!E{THEME[0] + i}", size=9, color=GREY_TXT, fmt=PCT, align=Alignment(horizontal="right"))
        put(f"Q{rr}", f"=Analisis!F{THEME[0] + i}", size=9, fmt='0.0"×"', align=Alignment(horizontal="right"))
    ws.conditional_formatting.add(f"N42:N{41 + len(themes)}",
                                  DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.45, color=ORANGE, showValue=True))
    put("K52", "Lift = % negatif ÷ % positif. Makin besar, makin khas penyebab kecewa.", merge_to="Q52", size=8, color=GREY_TXT, align=Alignment(indent=1))

    # --- Kategori & jenis kulit
    section_bar(55, "DI MANA PEMBELI KECEWA")
    card("B57:H73")
    card_title("B57", "ULASAN NEGATIF PER KATEGORI")
    for addr, t, al in [("B59", "Kategori", "left"), ("E59", "Ulasan", "right"), ("F59", "% negatif", "left")]:
        put(addr, t, size=8, bold=True, color=INK2, align=Alignment(horizontal=al, indent=1 if al == "left" else 0))
    ws.merge_cells("B59:D59"); ws.merge_cells("F59:H59")
    for col in "BCDEFGH":
        ws[f"{col}59"].border = Border(bottom=Side(style="thin", color="90A4AE"))
    for k in range(len(cat_order)):
        rr = 60 + k
        put(f"B{rr}", f"=Analisis!H{CAT[0] + k}", merge_to=f"D{rr}", size=9, align=Alignment(indent=1))
        put(f"E{rr}", f"=Analisis!K{CAT[0] + k}", size=9, fmt=INT, align=Alignment(horizontal="right"))
        put(f"F{rr}", f"=Analisis!L{CAT[0] + k}", merge_to=f"H{rr}", size=9, bold=True, fmt=PCT)
    ws.conditional_formatting.add(f"F60:F{59 + len(cat_order)}",
                                  DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.2, color=TEAL, showValue=True))
    ws.conditional_formatting.add(f"B60:H{59 + len(cat_order)}",
                                  FormulaRule(formula=['$B60="Sunscreen"'], fill=FILL("FDECE7"), font=Font(bold=True)))
    put(f"B{60 + len(cat_order)}", "Diurutkan dari % negatif tertinggi. Sunscreen disorot.", merge_to=f"H{60 + len(cat_order)}", size=8, color=GREY_TXT, align=Alignment(indent=1))

    card("I57:Q73")
    card_title("I57", "KELUHAN SUNSCREEN PER JENIS KULIT")
    put("I58", "% ulasan negatif yang menyebut tema. Tidak memakai filter jenis kulit.", merge_to="Q58", size=8, color=GREY_TXT, align=Alignment(indent=1))
    put("I60", "Tema keluhan", merge_to="L60", size=8, bold=True, color=INK2, align=Alignment(indent=1, vertical="center"))
    for j, sk in enumerate(SKINS[:4]):
        col = L(13 + j)
        put(f"{col}60", f'="{sk}"&CHAR(10)&"n = "&FIXED(Analisis!{L(2 + j)}{HEAT[0]},0)', size=8, bold=True, color=INK2,
            align=Alignment(horizontal="center", vertical="center", wrap_text=True))
    ws.row_dimensions[60].height = 26
    for col in "IJKLMNOP":
        ws[f"{col}60"].border = Border(bottom=Side(style="thin", color="90A4AE"))
    white = Side(style="thin", color="FFFFFF")
    for i in range(7):
        rr = 61 + i
        put(f"I{rr}", f"=Analisis!A{HEAT[0] + 1 + i}", merge_to=f"L{rr}", size=9, align=Alignment(indent=1))
        for j in range(4):
            c = put(f"{L(13 + j)}{rr}", f"=Analisis!{L(2 + j)}{HEAT[0] + 1 + i}", size=9, fmt=PCT0, align=Alignment(horizontal="center"))
            c.border = Border(left=white, right=white, top=white, bottom=white)
    ws.conditional_formatting.add("M61:P67", ColorScaleRule(start_type="min", start_color="F1FAF8", end_type="max", end_color="5BBFAF"))
    put("I69", "Kulit berminyak paling sering mengeluh tekstur berminyak; kulit kering lebih sering mengeluh iritasi.",
        merge_to="Q69", size=8, color=GREY_TXT, align=Alignment(indent=1, wrap_text=True, vertical="top"))

    # --- Produk
    section_bar(74, "PRODUK SUNSCREEN DENGAN KELUHAN TERTINGGI")
    card("B76:Q89")
    card_title("B76", "10 PRODUK DENGAN % ULASAN NEGATIF TERTINGGI (MIN. 500 ULASAN, SEMUA PERIODE)")
    heads = [("B78", "C78", "Brand", "left"), ("D78", "H78", "Produk", "left"), ("I78", "I78", "Ulasan", "right"),
             ("J78", "J78", "Rating", "right"), ("K78", "L78", "% negatif", "left"), ("M78", "N78", "% rekomendasi", "right"),
             ("O78", "Q78", "Keluhan utama", "left")]
    for a, b, t, al in heads:
        put(a, t, merge_to=(b if a != b else None), size=8, bold=True, color=INK2, align=Alignment(horizontal=al, indent=1 if al == "left" else 0))
    for col in "BCDEFGHIJKLMNOPQ":
        ws[f"{col}78"].border = Border(bottom=Side(style="thin", color="90A4AE"))
    for i in range(10):
        rr, src = 79 + i, H0 + 1 + i
        put(f"B{rr}", f"=Produk_Sunscreen!A{src}", merge_to=f"C{rr}", size=9, bold=True, align=Alignment(indent=1))
        put(f"D{rr}", f"=Produk_Sunscreen!B{src}", merge_to=f"H{rr}", size=9, color=INK2, align=Alignment(indent=1, shrink_to_fit=True))
        put(f"I{rr}", f"=Produk_Sunscreen!C{src}", size=9, fmt=INT, align=Alignment(horizontal="right"))
        put(f"J{rr}", f"=Produk_Sunscreen!H{src}", size=9, fmt=DEC2, align=Alignment(horizontal="right"))
        put(f"K{rr}", f"=Produk_Sunscreen!I{src}", merge_to=f"L{rr}", size=9, bold=True, fmt=PCT)
        put(f"M{rr}", f"=Produk_Sunscreen!J{src}", merge_to=f"N{rr}", size=9, fmt=PCT, align=Alignment(horizontal="right"))
        put(f"O{rr}", f'=Produk_Sunscreen!K{src}&"  ("&FIXED(Produk_Sunscreen!L{src}*100,0)&"%)"', merge_to=f"Q{rr}", size=9, align=Alignment(indent=1))
    ws.conditional_formatting.add("K79:K88", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=0.45, color=ORANGE, showValue=True))
    put("B89", "Daftar lengkap 20 produk ada di sheet Produk_Sunscreen. Rata-rata ulasan negatif sunscreen: 13,8%.",
        merge_to="Q89", size=8, color=GREY_TXT, align=Alignment(indent=1))

    # --- Temuan & sumber
    section_bar(92, "TEMUAN UTAMA (SEMUA DATA)")
    card("B94:Q101")
    notes = [
        "Sunscreen punya 13,8% ulasan negatif, tertinggi di antara kategori besar (rata-rata skincare 10,4%).",
        "Keluhan terbanyak adalah tekstur berminyak atau berat: 32% ulasan negatif sunscreen.",
        "Pilling 8,6× lebih sering di ulasan negatif dibanding ulasan positif.",
        "White cast dikeluhkan di 23% ulasan negatif dari kulit sawo matang, dibanding 11% dari kulit terang.",
        "Rekomendasi pengulas antar-kelompok harga hanya berbeda 4,3 poin (81,6% sampai 85,9%).",
        "Rekomendasi produk dan keterbatasan analisis ada di sheet Temuan.",
    ]
    for i, n in enumerate(notes):
        put(f"B{95 + i}", ("•  " if i < 5 else "") + n, merge_to=f"Q{95 + i}", size=9, color=(INK2 if i < 5 else GREY_TXT),
            italic=(i == 5), align=Alignment(indent=1))
    put("B103", "Sumber: Kaggle, Sephora Products and Skincare Reviews (nadyinky). Data pasar Amerika Serikat. "
                "Semua angka dihitung dengan rumus di sheet Analisis. Hubungan yang terlihat adalah korelasi.",
        merge_to="Q103", size=8, color=GREY_TXT, align=Alignment(indent=1))

    ws.print_area = "A1:R104"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ------------------------------------------------------------------ Temuan
    title_block(tm, "Temuan dan rekomendasi", "Berdasarkan seluruh data (tanpa filter). Angka bisa diperiksa di sheet Analisis dengan filter Dashboard = Semua.")
    tm.column_dimensions["A"].width = 6; tm.column_dimensions["B"].width = 38; tm.column_dimensions["C"].width = 90
    header(tm, 4, 1, ["No", "Temuan", "Penjelasan"])
    findings = [
        ("Harga bukan penentu kepuasan", "Korelasi harga dan rating hanya 0,13 (Spearman). Rekomendasi tertinggi di $50-100 (85,9%), produk di atas $100 justru turun (83,9%)."),
        ("Sunscreen kategori besar paling mengecewakan", "13,8% ulasan sunscreen negatif, 1,3× rata-rata skincare (10,4%). Self Tanners dan High Tech Tools lebih tinggi tetapi ulasannya jauh lebih sedikit."),
        ("Keluhan utama soal pengalaman pemakaian", "Berminyak/berat 32%, jerawat 15%, pilling 14%, white cast 13% dari ulasan negatif. Bau/wangi bukan penyebab kecewa (lift 0,7×)."),
        ("Pilling paling khas penyebab kecewa", "Pilling 8,6× lebih sering di ulasan negatif dibanding positif: jarang dipuji, langsung menurunkan rating."),
        ("Beda jenis kulit, beda keluhan", "Kulit berminyak: 41% mengeluh berminyak dan 19% jerawat. Kulit kering lebih sering mengeluh iritasi (14% vs 9%)."),
        ("White cast naik seiring warna kulit", "Terang 11%, medium 14%, sawo matang 23%, gelap 39%. Relevan untuk konsumen Indonesia yang mayoritas medium sampai sawo matang."),
    ]
    recs = [
        ("Tekstur ringan dan matte", "Menjawab keluhan nomor satu di semua jenis kulit, terutama kulit berminyak dan iklim tropis."),
        ("Tanpa white cast di kulit sawo matang", "Uji formula pada warna kulit lokal; komunikasikan hasilnya secara jelas."),
        ("Tidak menggumpal di bawah makeup", "Pilling paling khas ulasan negatif; uji dengan layering skincare dan makeup."),
        ("Non-comedogenic yang teruji", "Jerawat adalah keluhan kedua terbesar, tertinggi pada kulit berminyak."),
        ("Varian lembut untuk kulit kering/sensitif", "Kulit kering lebih sering mengeluh perih dan iritasi."),
    ]
    rr = 5
    for i, (a, b) in enumerate(findings, 1):
        tm.cell(row=rr, column=1, value=i).font = F()
        tm.cell(row=rr, column=2, value=a).font = F(bold=True)
        c = tm.cell(row=rr, column=3, value=b); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
        tm.row_dimensions[rr].height = 28
        rr += 1
    rr += 1
    tm.cell(row=rr, column=1, value="Rekomendasi untuk brand sunscreen lokal (contoh: Wardah/Emina), diurutkan menurut besarnya masalah").font = F(size=12, bold=True)
    rr += 1
    header(tm, rr, 1, ["No", "Rekomendasi", "Alasan dari data"])
    rr += 1
    for i, (a, b) in enumerate(recs, 1):
        tm.cell(row=rr, column=1, value=i).font = F()
        tm.cell(row=rr, column=2, value=a).font = F(bold=True)
        c = tm.cell(row=rr, column=3, value=b); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
        rr += 1
    rr += 1
    tm.cell(row=rr, column=1, value="Keterbatasan").font = F(size=12, bold=True)
    for i, t in enumerate(["Data Sephora pasar Amerika Serikat, bukan Indonesia; perlu divalidasi dengan ulasan lokal.",
                           "Tema keluhan dideteksi dengan kata kunci; masih ada salah tangkap.",
                           "Pengulas tidak mewakili semua pembeli (self-selection bias).",
                           "Hubungan yang terlihat adalah korelasi, bukan sebab-akibat."], 1):
        tm.cell(row=rr + i, column=2, value="•  " + t).font = F()

    # ------------------------------------------------------------------ Kualitas
    title_block(kq, "Pemeriksaan kualitas dan rekonsiliasi", "Membandingkan total di tabel agregat dengan angka dari pemrosesan Python (notebook 01). Status dihitung dengan rumus.")
    header(kq, 4, 1, ["Pemeriksaan", "Nilai di workbook", "Nilai acuan (Python)", "Status", "Catatan"])
    for c, w in zip("ABCDE", [46, 18, 20, 12, 70]):
        kq.column_dimensions[c].width = w
    n_sun_neg = int((sun["rating"] <= 2).sum()); n_sun_pos = int((sun["rating"] >= 4).sum())
    checks = [
        ("Total ulasan di Data_Ulasan", f"=SUM({UR('F')})", len(ru), "Ulasan dari 12 kategori ≥ 5.000 ulasan."),
        ("Ulasan yang dikeluarkan (kategori kecil)", f"={total_all}-B5", total_all - len(ru), "Kategori Shop by Concern (15 ulasan) tidak dianalisis."),
        ("Ulasan negatif di Data_Ulasan", f"=SUM({UR('H')})", int(ru["neg"].sum()), "Rating 1–2 bintang."),
        ("Ulasan sunscreen negatif di Data_Sunscreen", f'=SUMIFS({SR("E")},{SR("A")},"Negatif")', n_sun_neg, "Harus sama dengan ulasan sunscreen rating 1–2."),
        ("Ulasan sunscreen positif di Data_Sunscreen", f'=SUMIFS({SR("E")},{SR("A")},"Positif")', n_sun_pos, "Rating 4–5; rating 3 (netral) tidak dipakai di analisis tema."),
        ("Sunscreen: Data_Ulasan vs Data_Sunscreen", f'=SUMIFS({UR("F")},{UR("A")},"Sunscreen")', len(sun), "Ulasan sunscreen total (termasuk netral)."),
    ]
    for i, (a, f, ref, note) in enumerate(checks, start=5):
        kq.cell(row=i, column=1, value=a).font = F()
        kq.cell(row=i, column=2, value=f).number_format = INT
        kq.cell(row=i, column=3, value=ref).number_format = INT
        kq.cell(row=i, column=4, value=f'=IF(B{i}=C{i},"Lolos","Periksa")')
        kq.cell(row=i, column=5, value=note).font = F(color=GREY_TXT)
        for j in (2, 3, 4):
            kq.cell(row=i, column=j).font = F()
    kq.cell(row=12, column=1, value="Catatan: nilai acuan diketik dari hasil Python saat workbook dibuat (sumber: notebooks/01_data_understanding_cleaning.ipynb).").font = F(color=GREY_TXT)

    # ------------------------------------------------------------------ Kamus_Data
    title_block(kd, "Kamus data", "Arti setiap kolom di sheet data.")
    header(kd, 4, 1, ["Sheet", "Kolom", "Arti", "Catatan"])
    for c, w in zip("ABCD", [18, 26, 60, 60]):
        kd.column_dimensions[c].width = w
    dict_rows = [
        ("Data_Ulasan", "Kategori", "Kategori sekunder produk skincare Sephora", "12 kategori dengan ≥ 5.000 ulasan"),
        ("Data_Ulasan", "Kelompok harga", "Harga produk (USD) dalam 4 kelompok", "< $25, $25-50, $50-100, > $100"),
        ("Data_Ulasan", "Jenis kulit", "Jenis kulit yang diisi pengulas", "Kosong diganti ‘Tidak diketahui’"),
        ("Data_Ulasan", "Warna kulit", "Warna kulit pengulas, dikelompokkan", "Terang = fair/fairLight/light; Medium = lightMedium/medium; Sawo matang = mediumTan/tan; Gelap = deep/rich/dark"),
        ("Data_Ulasan", "Tahun", "Tahun ulasan dikirim", "2008–2023; 2023 hanya Januari–Maret"),
        ("Data_Ulasan", "Jumlah ulasan", "Banyaknya ulasan pada kombinasi tersebut", ""),
        ("Data_Ulasan", "Total bintang", "Jumlah rating bintang (1–5)", "Rating rata-rata = Total bintang ÷ Jumlah ulasan"),
        ("Data_Ulasan", "Ulasan negatif (1–2★)", "Ulasan dengan rating 1 atau 2", ""),
        ("Data_Ulasan", "Rekomendasi: ya / dijawab", "Jawaban ‘apakah Anda merekomendasikan’", "% merekomendasikan = ya ÷ dijawab. Pertanyaan ini baru ada sejak 2017 (sekitar separuh ulasan 2017, semua ulasan mulai 2018); ulasan 2008–2016 tidak dihitung."),
        ("Data_Sunscreen", "Sentimen", "Negatif = rating 1–2, Positif = rating 4–5", "Rating 3 tidak dipakai"),
        ("Data_Sunscreen", "Kolom tema", "Jumlah ulasan yang mengeluhkan tema", "Kata kunci tanpa negasi dalam 3–4 kata sebelumnya; lihat notebook 03"),
        ("Produk_Sunscreen", "Keluhan utama", "Tema paling sering di ulasan negatif produk", "Dihitung di Python"),
    ]
    for i, row_ in enumerate(dict_rows, start=5):
        for j, v in enumerate(row_, start=1):
            c = kd.cell(row=i, column=j, value=v); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
    add_table(kd, "TblKamus", f"A4:D{4 + len(dict_rows)}")

    # ------------------------------------------------------------------ Dokumentasi
    title_block(doc, "Dokumentasi proyek", "Cara memakai, alur kerja, dan batasan.")
    doc.column_dimensions["A"].width = 28; doc.column_dimensions["B"].width = 110
    header(doc, 4, 1, ["Topik", "Penjelasan"])
    docs = [
        ("Tujuan", "Menemukan kategori skincare yang paling mengecewakan dan keluhan utama pembeli sunscreen, sebagai masukan produk untuk brand lokal."),
        ("Sumber data", "Kaggle: Sephora Products and Skincare Reviews (nadyinky). 2.420 produk skincare, 1.092.967 ulasan, Agustus 2008 – Maret 2023."),
        ("Cara memakai", "Buka Dashboard, ubah sel kuning (jenis kulit, warna kulit, tahun). Semua KPI, grafik, dan heatmap dihitung ulang. Telusuri angka ke sheet Analisis."),
        ("Alur kerja", "Python (pandas): gabung 5 file ulasan, bersihkan, deteksi tema keluhan → ringkas menjadi Data_Ulasan dan Data_Sunscreen → Excel menghitung semua metrik dengan SUMIFS."),
        ("Kenapa data diringkas", "Ulasan mentah 1,09 juta baris melebihi batas Excel (1.048.576 baris). Agregasi tidak mengubah hasil karena semua metrik berupa jumlah dan rasio jumlah."),
        ("Rasio", "Selalu dihitung dari total pembilang ÷ total penyebut, bukan rata-rata persentase kelompok."),
        ("Filter khusus", "Grafik white cast tidak memakai filter warna kulit; heatmap tidak memakai filter jenis kulit, agar kelompok bisa dibandingkan."),
        ("Memperbarui", "Jalankan notebook 01, lalu src/build_excel.py, lalu buka di Excel (atau src/recalc_excel.ps1) agar rumus dihitung."),
        ("File terkait", "notebooks/01–03 (analisis), dashboard/index.html (versi web), reports/figures (grafik statis)."),
    ]
    for i, (a, b) in enumerate(docs, start=5):
        doc.cell(row=i, column=1, value=a).font = F(bold=True)
        c = doc.cell(row=i, column=2, value=b); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")

    for wsx in (an, kq, kd, doc, tm):
        wsx.sheet_view.showGridLines = False
    dash.sheet_properties.tabColor = NAVY
    wb.active = 0
    wb.save(OUT)
    print("Tersimpan:", OUT)
    print("Baris Data_Ulasan:", len(cube), "| Data_Sunscreen:", len(scube), "| Produk:", len(prod))


if __name__ == "__main__":
    main()
