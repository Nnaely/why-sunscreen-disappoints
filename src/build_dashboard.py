"""Bangun dashboard offline: agregasi data -> tanam JSON ke dashboard/template.html -> dashboard/index.html.

Jalankan dari mana saja:  .venv/Scripts/python.exe src/build_dashboard.py
Butuh: data/processed/reviews_clean.parquet (hasil notebook 01).
"""
from pathlib import Path
import json
import re

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
PROCESSED = ROOT / "data" / "processed"
DASH = ROOT / "dashboard"

# Tema keluhan: sama persis dengan notebooks/03_sunscreen_complaints.ipynb
THEMES = {
    "Berminyak / lengket / berat": r"greas|oily(?!\s*(?:skin|and|/|-|,|t.zone|combo|or|prone|complexion|to))|shiny|\bshine\b|sticky|tacky|slick|heavy",
    "Jerawat / breakout": r"break\s?out|broke (?:me )?out|breaking out|\bacne\b(?![\s-]*prone)|pimple|clog|cystic|whitehead|milia|\bbumps\b",
    "Pilling (menggumpal)": r"\bpill|\bpils\b|balls? up|rolls? off|crumbl",
    "White cast": r"white[\s-]?cast|whitecast|chalky|ashy|ghostly|ghost|purple cast|gr[ae]y cast|white residue|leaves? (?:a |me )?(?:slight )?white",
    "Iritasi / perih": r"\bburn(?:ing|ed|s)?\b|\bsting|\bitch|\brash|redness|irritat|allerg|react(?:ion|ed)",
    "Bikin kering": r"\bdry(?:ing)?\b(?!\s*(?:skin|/|-|,|and|or|to|combo|sensitive))|flaky|flakiness|tight|dehydrat|parch",
    "Bau / wangi": r"\bsmell|scent|fragranc|odor|odour|perfum|stink",
    "Harga / ukuran": r"overpriced|expensive|pricey|waste of money|\bprice\b|tiny (?:bottle|tube)|small (?:bottle|tube)",
    "Perih di mata": r"(?:eye|eyes)\W+(?:\w+\W+){0,4}(?:sting|burn|water|tear)|(?:sting|burn)\w*\W+(?:\w+\W+){0,4}eyes",
    "Tidak melindungi": r"sunburn|sun burn|got burn|got burnt|still burn|burned me|didn.t protect|not protect",
}
NEGATION = re.compile(r"(?:\bno\b|\bnot\b|n't|n’t|\bwithout\b|\bzero\b|\bnever\b|\bnor\b|\bnon\b|\blittle\b|\bminimal\b|\bany\b)\W+(?:\w+\W+){0,3}$")

TONE_MAP = {
    "fair": "Terang", "fairLight": "Terang", "light": "Terang",
    "lightMedium": "Medium", "medium": "Medium",
    "mediumTan": "Sawo matang", "tan": "Sawo matang",
    "deep": "Gelap", "rich": "Gelap", "dark": "Gelap",
}
TONES = ["Terang", "Medium", "Sawo matang", "Gelap", "Tidak diketahui"]
SKINS = ["combination", "dry", "normal", "oily", "unknown"]
PRICE_BANDS = ["< $25", "$25-50", "$50-100", "> $100"]


def is_complaint(text, pattern):
    for m in re.finditer(pattern, text):
        if not NEGATION.search(text[max(0, m.start() - 40):m.start()]):
            return True
    return False


def main():
    r = pd.read_parquet(PROCESSED / "reviews_clean.parquet",
                        columns=["rating", "is_recommended", "skin_type", "skin_tone", "year",
                                 "secondary_category", "price_band", "product_name", "brand_name",
                                 "review_title", "review_text"])
    r["tone"] = r["skin_tone"].map(TONE_MAP).fillna("Tidak diketahui")
    r["price_band"] = r["price_band"].astype(str)
    r["neg"] = (r["rating"] <= 2).astype(int)
    r["pos"] = (r["rating"] >= 4).astype(int)
    r["rec_yes"] = (r["is_recommended"] == 1).astype(int)
    r["rec_n"] = r["is_recommended"].notna().astype(int)

    cats = (r["secondary_category"].value_counts().loc[lambda s: s >= 5000].index.tolist())
    years = sorted(r["year"].unique().tolist())

    # Kubus utama: kategori x harga x jenis kulit x warna kulit x tahun
    dims = ["secondary_category", "price_band", "skin_type", "tone", "year"]
    cube = (r[r["secondary_category"].isin(cats)]
            .groupby(dims, observed=True)
            .agg(n=("rating", "size"), rating_sum=("rating", "sum"), neg=("neg", "sum"),
                 rec_yes=("rec_yes", "sum"), rec_n=("rec_n", "sum"))
            .reset_index())
    cube_rows = [[cats.index(c), PRICE_BANDS.index(p), SKINS.index(s), TONES.index(t), int(y),
                  int(n), int(rs), int(ng), int(ry), int(rn)]
                 for c, p, s, t, y, n, rs, ng, ry, rn in cube.itertuples(index=False)]

    # Sunscreen: tema keluhan per kelompok sentimen
    sun = r[r["secondary_category"] == "Sunscreen"].copy()
    text = (sun["review_title"].fillna("") + ". " + sun["review_text"].fillna("")).str.lower()
    theme_names = list(THEMES)
    for t in theme_names:
        sun[t] = text.map(lambda x, p=THEMES[t]: is_complaint(x, p)).astype(int)
    sun["grp"] = np.select([sun["rating"] <= 2, sun["rating"] >= 4], [0, 1], 2)  # 0=neg 1=pos 2=netral
    sun = sun[sun["grp"] < 2]
    sc = sun.groupby(["grp", "skin_type", "tone", "year"]).agg(n=("rating", "size"), **{t: (t, "sum") for t in theme_names}).reset_index()
    # Baris: [grp, skin, tone, year, n, jumlah_tema_1 ... jumlah_tema_10]
    sun_rows = [[int(g), SKINS.index(s), TONES.index(t), int(y), int(n)] + [int(v) for v in rest]
                for g, s, t, y, n, *rest in sc.itertuples(index=False, name=None)]

    # Tabel produk sunscreen (>= 500 ulasan, semua periode)
    sall = r[r["secondary_category"] == "Sunscreen"]
    prod = (sall.groupby(["brand_name", "product_name"])
            .agg(n=("rating", "size"), rating=("rating", "mean"), neg=("neg", "mean"),
                 rec=("rec_yes", "sum"), rec_n=("rec_n", "sum"))
            .query("n >= 500").reset_index())
    negs = sun[sun["grp"] == 0].groupby("product_name")[theme_names].mean()
    products = []
    for p in prod.itertuples(index=False):
        top = negs.loc[p.product_name].sort_values(ascending=False) if p.product_name in negs.index else None
        products.append({
            "brand": p.brand_name, "product": p.product_name, "n": int(p.n),
            "rating": round(float(p.rating), 2), "neg": round(float(p.neg), 4),
            "rec": round(p.rec / p.rec_n, 4) if p.rec_n else None,
            "issue": top.index[0] if top is not None and top.iloc[0] > 0 else "-",
            "issue_pct": round(float(top.iloc[0]), 3) if top is not None else None,
        })

    data = {
        "categories": cats, "priceBands": PRICE_BANDS, "skins": SKINS, "tones": TONES,
        "years": [int(y) for y in years], "themes": theme_names,
        "cube": cube_rows, "sun": sun_rows, "products": products,
        "meta": {"total_reviews": int(len(r)), "period": f"{r['year'].min()}–{r['year'].max()}"},
    }

    # Rekonsiliasi: total kubus harus sama dengan jumlah ulasan kategori yang dipakai
    assert sum(row[5] for row in cube_rows) == int(r["secondary_category"].isin(cats).sum())
    assert sum(row[4] for row in sun_rows) == len(sun)

    tpl = (DASH / "template.html").read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = tpl.replace("/*__DATA__*/null", payload)
    (DASH / "index.html").write_text(html, encoding="utf-8")
    print(f"OK: {len(cube_rows):,} baris kubus, {len(sun_rows):,} baris sunscreen, {len(products)} produk "
          f"-> {DASH / 'index.html'} ({len(html) / 1024:,.0f} KB)")


if __name__ == "__main__":
    main()
