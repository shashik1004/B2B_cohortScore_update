"""
Build the executive deck for the B2B client score analysis (TotalScore basis).

  methodology_slides — overview (what / why / how) and cohort membership & definitions.
  results_slides     — score impact (cross-tab, every client, top 25 lenders);
                       why scores move (benchmark + inside the calculation);
                       stability (periods loaded from the analysis configuration).

Python port of deck_style_kit/deck_style.js (same palette, fonts, coordinates).

Sources
  b2b_stability_outputs/b2b_client_score_stability_results.xlsx
      (B2B_Client_Score_Stability_Analysis.ipynb)
  B2B_cohort_score_20260531_upd.xlsx  - rawTotal + cprAuditLong (Cohort approach)
  B2B_reporting_raw_20260531.xlsx     - rawTotal (Old approach)

    python build_stability_slide.py   ->  B2B_Score_Stability_Exec_Slide.pptx
"""
import ast
import re
import sys
from pathlib import Path

import numpy as np

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------- palette / fonts (from deck_style.js) ----------
INDIGO, INDIGO_DK, INDIGO_TINT = "2F20EF", "160C63", "ECEBFE"
RED, RED_TINT = "C8102E", "FBE9EC"
GREEN, GREEN_TINT = "1B8A5A", "E7F4EE"
AMBER, AMBER_TINT = "B26B00", "FBEFD8"
INK, MUTE, GRAYFILL, LINEGRAY, WHITE = "1C1C28", "6B7280", "F4F5F9", "E3E5EE", "FFFFFF"
INDIGO_TEXT_LT, INDIGO_TEXT_MUTE, GOLD = "CFCBFB", "9C93F5", "FFD98A"
REFGRAY = "C2C7DA"
F, FH = "Calibri", "Cambria"
W, H = 13.33, 7.5
ORG_FOOTER = "© 2026 Onity Group Inc. All rights reserved. Proprietary and Confidential"

ROOT = Path(__file__).resolve().parent
RESULTS = ROOT.parent / "b2b_stability_outputs" / "b2b_client_score_stability_results.xlsx"
COHORT_FILE = ROOT /  "B2B_cohort_score_20260531_upd.xlsx"
OLD_FILE = ROOT / "B2B_reporting_raw_20260531.xlsx"
OUT = ROOT /  "B2B_Score_Stability_Exec_Slide.pptx"
STYLE_KIT = (ROOT.parents[3] / "00_AI_Agentic_Dev" / "dsa-agent-harness-prototype"
             / "skills" / "exec-deck" / "scripts")
sys.path.insert(0, str(STYLE_KIT))
from deck_style import new_deck, render_png


def rgb(hex_):
    return RGBColor.from_string(hex_)


# ---------- primitive helpers ----------
def box(slide, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None, shape=None):
    assert w > 0 and h > 0, f"non-positive box {w}x{h} at ({x},{y})"
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        shp.adjustments[0] = min(0.5, radius / min(w, h))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def line(slide, x1, y1, x2, y2, color, width=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


def text(slide, x, y, w, h, content, size=12, color=INK, bold=False, italic=False,
         font=F, align="left", valign="top", spacing=None, char_spacing=None):
    """content: str, or list of paragraphs; a paragraph is str or list of (text, overrides)."""
    assert w > 0 and h > 0, f"non-positive text box {w}x{h} at ({x},{y})"
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[valign]
    paras = [content] if isinstance(content, str) else content
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT}[align]
        if spacing:
            p.line_spacing = spacing
        runs = [(para, {})] if isinstance(para, str) else para
        for txt, o in runs:
            r = p.add_run()
            r.text = txt
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.color.rgb = rgb(o.get("color", color))
            cs = o.get("char_spacing", char_spacing)
            if cs:
                r._r.get_or_add_rPr().set("spc", str(int(cs * 100)))
        if isinstance(para, list) and para and para[0][1].get("space_before"):
            p.space_before = Pt(para[0][1]["space_before"])
    return tb


# ---------- house-style helpers (ports of deck_style.js) ----------
def footer(slide, n):
    spec = STRETCH.pop(slide.slide_id, None)
    if spec:
        stretch_content(slide, *spec)
    bar_y, bar_h = 7.02, 0.48
    box(slide, 0, bar_y, W, bar_h, fill=INDIGO)
    text(slide, 0.45, bar_y, W - 1.55, bar_h, ORG_FOOTER, size=9, color="D8D5FB", valign="middle")
    text(slide, W - 1.05, bar_y, 0.6, bar_h, str(n), size=10, color=WHITE, bold=True,
         align="right", valign="middle")


def tag(slide, label, color):
    tw = 0.3 + len(label) * 0.095
    box(slide, W - 0.45 - tw, 0.42, tw, 0.34, fill=color, radius=0.17)
    text(slide, W - 0.45 - tw, 0.42, tw, 0.34, label, size=10, bold=True, color=WHITE,
         align="center", valign="middle", char_spacing=1)


def slide_title(slide, title, sub=None):
    text(slide, 0.5, 0.44, 10.5, 0.95, title, size=23, bold=True, font=FH, spacing=0.96)
    if sub:
        text(slide, 0.5, 1.32, 12.2, 0.5, sub, size=13, italic=True, color=MUTE, spacing=0.98)


def panel_header(slide, x, y, w, num, eyebrow, header, accent=INDIGO):
    d = 0.3
    box(slide, x, y + 0.02, d, d, fill=accent, radius=d / 2)
    text(slide, x, y + 0.02, d, d, str(num), size=11, bold=True, color=WHITE,
         align="center", valign="middle")
    text(slide, x + 0.42, y + 0.04, w - 0.42, 0.26, eyebrow, size=10, bold=True,
         color=accent, char_spacing=1.5, valign="middle")
    text(slide, x, y + 0.42, w, 0.56, header, size=14, bold=True, font=FH, spacing=0.95)


# Compact deck layout retains the existing panels and reserves a slim takeaway band.
DECK_MODE = {"compact": False}
CONTENT_BOTTOM = 6.24
STRETCH = {}          # slide_id -> (y0, old content bottom, new content bottom); applied by footer()


def stretch_content(slide, y0, y1_old, y1_new, font_k=1.02):
    """Scale the content zone [y0, y1_old] down to y1_new: positions and heights stretch, squares and
    circles keep their shape, and fonts grow slightly, so charts, tables and cards read larger."""
    from pptx.shapes.connector import Connector
    E = 914400
    k = (y1_new - y0) / (y1_old - y0)
    a0, b0 = y0 * E, y1_old * E
    f = lambda v: int(a0 + (v - a0) * k)
    for shp in list(slide.shapes):
        if isinstance(shp, Connector):
            ys = [shp.begin_y, shp.end_y]
            if min(ys) >= a0 - 0.01 * E and max(ys) <= b0 + 0.02 * E:
                by_, ey_ = f(shp.begin_y), f(shp.end_y)
                shp.begin_y, shp.end_y = by_, ey_
            continue
        top, h = shp.top, shp.height
        if top < a0 - 0.01 * E or top + h > b0 + 0.02 * E:
            continue
        if abs(shp.width - h) < 0.03 * E:
            shp.top = int(f(top + h / 2) - h / 2)
        else:
            shp.top, shp.height = f(top), int(h * k)
        if shp.has_text_frame:
            for para in shp.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size:
                        r.font.size = Pt(round(r.font.size.pt * font_k * 2) / 2)


def impact_band(slide, label, runs, y=6.09, h=0.8, size=12):
    if DECK_MODE["compact"]:
        y, h, size = 6.38, 0.51, 10.5
    box(slide, 0.5, y, W - 1.0, h, fill=INDIGO_DK, radius=0.08)
    text(slide, 0.75, y, 1.55, h, label, size=10.5, bold=True, color=INDIGO_TEXT_MUTE,
         char_spacing=1.5, valign="middle")
    text(slide, 2.35, y, W - 0.5 - 2.35 - 0.25, h, [runs], size=size, color=INDIGO_TEXT_LT,
         valign="middle", spacing=1.02)


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if DECK_MODE["compact"]:
        STRETCH[s.slide_id] = (PY, 5.95, CONTENT_BOTTOM)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(WHITE)
    return s


def pct(v, d=0):
    return f"{v * 100:.{d}f}%"


# ---------- load analysis numbers ----------
def read_xl(path, **kw):
    """pd.read_excel that falls back to a temp copy when the workbook is open (locked) in Excel."""
    try:
        return pd.read_excel(path, **kw)
    except PermissionError:
        import shutil, tempfile
        tmp = Path(tempfile.gettempdir()) / f"_deckread_{Path(path).name}"
        shutil.copy2(path, tmp)
        return pd.read_excel(tmp, **kw)


def notebook_config():
    """Tag rule and forward-test periods, read from the notebook's CONFIG cell so they never drift."""
    import json
    nb = json.loads((ROOT.parent / "B2B_Client_Score_Stability_Analysis.ipynb").read_text(encoding="utf-8"))
    src = next("".join(c["source"]) for c in nb["cells"] if "CONFIG = {" in "".join(c["source"]))
    assignment = next(node for node in ast.parse(src).body if isinstance(node, ast.Assign)
                      and any(isinstance(target, ast.Name) and target.id == "CONFIG" for target in node.targets))
    config = ast.literal_eval(assignment.value)
    return {"fav_mult": config["favourable_multiplier"], "unf_mult": config["unfavourable_multiplier"],
            "score_column": config["score_column"], "base_period": config["forward_periods"]["Base"],
            "comp_period": config["forward_periods"]["Comparison"], "files": config["files"],
            "forward_files": config["forward_files"]}


NB = notebook_config()
assert NB["score_column"] == "TotalScore" and NB["unf_mult"] < NB["fav_mult"], NB
COHORT_FILE = ROOT.parent / NB["files"]["Cohort Base"]
OLD_FILE = ROOT.parent / NB["files"]["Old Base"]
# Tags are relative to each population's own portfolio score Z: Favourable >= FAV_T x Z, Unfavourable < UNF_T x Z
FAV_T, UNF_T = NB["fav_mult"], NB["unf_mult"]
RULE = f"Favourable ≥ {FAV_T:g}×, Unfavourable < {UNF_T:g}× the portfolio score"
_MON = {m: i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}


def period_parts(p):
    """'October 2025-January 2026' -> ('Oct', 'Jan', 'Oct 2025 – Jan 2026', n_months)."""
    ms = re.findall(r"(January|February|March|April|May|June|July|August|September|October|November|December)", p)
    ys = re.findall(r"20\d\d", p)
    a, b = ms[0][:3], ms[-1][:3]
    n = (_MON[b] - _MON[a]) % 12 + 1
    long_ = f"{a} {ys[0]} – {b} {ys[-1]}" if len(ys) > 1 else f"{a} – {b} {ys[0]}"
    return a, b, long_, n


FB = period_parts(NB["base_period"])        # forward-test base window
FC = period_parts(NB["comp_period"])        # forward-test comparison window
F_BASE, F_COMP = f"{FB[0]}–{FB[1]}", f"{FC[0]}–{FC[1]}"
N_MONTHS_WORD = {3: "three", 4: "four", 5: "five", 6: "six"}.get(FC[3], str(FC[3]))

APPROACHES = ["Cohort", "Old"]
BANDS = ["Favourable", "Neutral", "Unfavourable"]
OLD_C, OLD_TINT = "5F6673", "ECEEF2"      # Old approach: grey (Cohort stays indigo)

x = read_xl(RESULTS, sheet_name=None)
thr = x["Thresholds"].set_index("Dataset")
assert thr["Score column"].eq(NB["score_column"]).all()
assert np.allclose(thr["Upper multiplier"], FAV_T) and np.allclose(thr["Lower multiplier"], UNF_T)
quality = x["Data Quality"].set_index("Dataset")
FUND_BEGIN = pd.to_datetime(quality.loc["Cohort Base", "Begin date"])
FUND_END = pd.to_datetime(quality.loc["Cohort Base", "End date"])
assert pd.to_datetime(quality["Begin date"]).eq(FUND_BEGIN).all()
assert pd.to_datetime(quality["End date"]).eq(FUND_END).all()
FUND_WINDOW = f"{FUND_BEGIN:%b %Y} – {FUND_END:%b %Y}"
forward_thresholds = x["Forward Thresholds"].set_index("Dataset")
for dataset, source in NB["forward_files"].items():
    expected_period = NB["base_period"] if dataset.endswith("Base") else NB["comp_period"]
    row = forward_thresholds.loc[dataset]
    assert row["Source file"] == source and row["Period"] == expected_period
    assert row["Score column"] == NB["score_column"]
    assert np.isclose(row["Upper multiplier"], FAV_T) and np.isclose(row["Lower multiplier"], UNF_T)
Z = {"Cohort": float(thr.loc["Cohort Base", "Z"]), "Old": float(thr.loc["Old Base", "Z"])}
cmp_counts = x["Base Compare Counts"]
xtab = (cmp_counts.set_index("OldBand") if "OldBand" in cmp_counts
    else cmp_counts.set_index("CohortBand").T).reindex(index=BANDS, columns=BANDS)
correct = x["Correct Identification"].set_index(["Approach", "Outcome"])   # forward test
next_tr = {a: x[f"{a} Next Transitions"].set_index("BaseBand") for a in APPROACHES}

n_clients = int(xtab.values.sum())
same = sum(int(xtab.loc[b, b]) for b in BANDS)
opposite = int(xtab.loc["Favourable", "Unfavourable"] + xtab.loc["Unfavourable", "Favourable"])
one_step = n_clients - same - opposite
up = sum(int(xtab.iloc[i, j]) for i in range(3) for j in range(3) if j < i)     # better tag under Cohort
down = sum(int(xtab.iloc[i, j]) for i in range(3) for j in range(3) if j > i)
fav_total = {"Cohort": int(xtab["Favourable"].sum()), "Old": int(xtab.loc["Favourable"].sum())}
unf_total = {"Cohort": int(xtab["Unfavourable"].sum()), "Old": int(xtab.loc["Unfavourable"].sum())}
unf_to_fav = int(xtab.loc["Unfavourable", "Favourable"])
fav_to_unf = int(xtab.loc["Favourable", "Unfavourable"])


hit = {(a, b): int(correct.loc[(a, f"Correctly identified {b}"), "Correct count"])
       for a in APPROACHES for b in ["Favourable", "Unfavourable"]}
tagged = {(a, b): int(correct.loc[(a, f"Correctly identified {b}"), "Base tagged count"])
          for a in APPROACHES for b in ["Favourable", "Unfavourable"]}
n_matched = int(correct["Total matched clients"].iloc[0])
hit_rate = {k: hit[k] / tagged[k] for k in hit}
total_hits = {a: hit[(a, "Favourable")] + hit[(a, "Unfavourable")] for a in APPROACHES}
total_tagged = {a: tagged[(a, "Favourable")] + tagged[(a, "Unfavourable")] for a in APPROACHES}
tag_rate = {a: total_hits[a] / total_tagged[a] for a in APPROACHES}          # share of tags confirmed
good_flagged_bad = {a: int(next_tr[a].loc["Unfavourable", "Favourable"]) for a in APPROACHES}


def client_rows(path, sheet="rawTotal"):
    raw = read_xl(path, sheet_name=sheet)
    tot = raw[raw["MetricName"].eq("Total") & raw["MetricValue"].eq("Total")]
    if "ProductCategory" in raw:
        tot = tot[tot["ProductCategory"].eq("All")] if sheet == "rawTotal" else tot
    return raw[raw["MetricName"].eq("Client")].set_index("MetricValue"), tot.iloc[0]


cohort_clients, cohort_total = client_rows(COHORT_FILE)
old_clients, old_total = client_rows(OLD_FILE)
for clients in (cohort_clients, old_clients):
    assert pd.to_datetime(clients["BeginFundDt"]).eq(FUND_BEGIN).all()
    assert pd.to_datetime(clients["EndFundDt"]).eq(FUND_END).all()
assert abs(cohort_total["TotalScore"] - Z["Cohort"]) < 1e-9 and abs(old_total["TotalScore"] - Z["Old"]) < 1e-9
audit = read_xl(COHORT_FILE, sheet_name="cprAuditLong")
inc = audit[audit["Included_Flag"].eq(1)]


def band_of(s, z=1.0):
    """Notebook rule: Favourable >= FAV_T x Z; Unfavourable strictly < UNF_T x Z."""
    s = pd.Series(s)
    return pd.Series(np.select([s >= FAV_T * z, s < UNF_T * z], ["Favourable", "Unfavourable"], "Neutral"),
                     index=s.index)


scores = pd.DataFrame({
    "rank": cohort_clients["VolumeRank"], "funded": cohort_clients["FundedLoanAmt"],
    "new": cohort_clients["TotalScore"], "old": old_clients["TotalScore"].reindex(cohort_clients.index),
})
assert scores[["new", "old"]].notna().all().all()
scores["new_rel"], scores["old_rel"] = scores["new"] / Z["Cohort"], scores["old"] / Z["Old"]
scores["new_band"], scores["old_band"] = band_of(scores["new"], Z["Cohort"]), band_of(scores["old"], Z["Old"])
# tie-out: re-banded source scores reproduce the notebook's cross-tab
assert (pd.crosstab(scores["old_band"], scores["new_band"]).reindex(index=BANDS, columns=BANDS, fill_value=0)
        .values == xtab.values).all()
bi = {b: i for i, b in enumerate(BANDS)}
scores["move"] = np.sign(scores["old_band"].map(bi) - scores["new_band"].map(bi))   # +1 better under Cohort
# clients with no prepayments in the window: the Cohort file assigns them a default TotalScore
scores["zero_prepay"] = cohort_clients["CPRScore"].eq(0) & cohort_clients["Client_ActualCPR"].eq(0)
zero = scores[scores["zero_prepay"]]
zero_default = float(zero["new"].round(6).mode().iloc[0]) if len(zero) else None
zero_unf = int(zero["new_band"].eq("Unfavourable").sum())
share_higher = (scores["new_rel"] > scores["old_rel"]).mean()
top25 = scores.sort_values("rank").head(25)
top25_share = top25["funded"].sum() / scores["funded"].sum()
top25_moves = {k: int((top25["move"] == v).sum()) for k, v in [("up", 1), ("same", 0), ("down", -1)]}
top25_unf = {"Old": int(top25["old_band"].eq("Unfavourable").sum()),
             "Cohort": int(top25["new_band"].eq("Unfavourable").sum())}
tag_share = {a: scores[f"{k}_band"].value_counts(normalize=True).reindex(BANDS, fill_value=0)
             for a, k in [("Old", "old"), ("Cohort", "new")]}

# severe reversals (notebook cell 21): both directions, with each file's components
rev = {"UF": x["Cohort Fav Old Unfav"], "FU": x["Cohort Unfav Old Fav"]}     # UF = Old U -> Cohort F
assert len(rev["UF"]) == unf_to_fav and len(rev["FU"]) == fav_to_unf
rev_zero = {k: int(v["MetricValue"].map(scores["zero_prepay"]).sum()) for k, v in rev.items()}
fwd_rev = {w: (len(x[f"{w} Cohort F Old U"]), len(x[f"{w} Cohort U Old F"])) for w in [F_BASE.replace("–", "-"),
                                                                                       F_COMP.replace("–", "-")]
           if f"{w} Cohort F Old U" in x}

# benchmark comparison: Old model CPR vs Cohort benchmark, client by client and by loan type
bench = pd.DataFrame({"old": old_clients["ModelCPR"], "new": cohort_clients["Cohort_ActualCPR"].reindex(old_clients.index),
                      "w": cohort_clients["FundedLoanAmt"].reindex(old_clients.index)}).dropna()
bench_n, bench_lower = len(bench), int((bench["old"] < bench["new"]).sum())
bench_lower_vol = bench.loc[bench["old"] < bench["new"], "w"].sum() / bench["w"].sum()
bench_lt = []
for sheet, label in [("rawConv", "Conventional"), ("rawFHA", "FHA"), ("rawVA", "VA")]:
    _, ot = client_rows(OLD_FILE, sheet)
    _, ct = client_rows(COHORT_FILE, sheet)
    bench_lt.append((label, ot["ModelCPR"], ct["Cohort_ActualCPR"]))
bench_exceptions = [(lab, o_, c_) for lab, o_, c_ in bench_lt if o_ > c_]
worst_lt = max(bench_lt, key=lambda r: r[2] / r[1])

# deep-dive clients: largest Old-Unfavourable -> Cohort-Favourable, and the largest client Cohort tags lower
real = scores[~scores["zero_prepay"] & cohort_clients["Client_ActualCPR"].notna()]
flip_leader = real[real["old_band"].eq("Unfavourable") & real["new_band"].eq("Favourable")]["rank"].idxmin()
down_leader = real[real["move"].eq(-1)]["rank"].idxmin()
DEEP_DIVE_CLIENTS = [flip_leader, down_leader]


def ratio_words(r):
    return ("less than a third as many" if r < 1 / 3 else "about half as many" if r <= 0.6 else
            "fewer" if r < 0.9 else "about as many")


# slide-5 caption: say plainly, band by band, which approach's tags hold better
_hr = lambda a, b: pct(hit_rate[(a, b)])
_gap = lambda b: hit_rate[("Cohort", b)] - hit_rate[("Old", b)]


def _band_phrase(b):
    g = _gap(b)
    if abs(g) < 0.02:
        return f"{b} tags hold about equally ({_hr('Cohort', b)} vs {_hr('Old', b)})"
    adv = "slightly more often" if abs(g) < 0.05 else "more often" if abs(g) < 0.1 else "far more often"
    who = "Cohort" if g > 0 else "Old"
    return f"{b} tags hold {adv} under {who} ({_hr('Cohort', b)} vs {_hr('Old', b)})"


_order = sorted(["Favourable", "Unfavourable"], key=lambda b: -abs(_gap(b)))
caption_lead = f"{_band_phrase(_order[0])}; {_band_phrase(_order[1])}."
caption_lead = caption_lead[0].upper() + caption_lead[1:]
caption_notes = []
for b in ["Favourable", "Unfavourable"]:
    if hit[("Old", b)] > hit[("Cohort", b)]:
        caption_notes.append(f"Old confirms {hit[('Old', b)] - hit[('Cohort', b)]} more {b} clients in count "
                             f"({hit[('Old', b)]} vs {hit[('Cohort', b)]}) because it tags "
                             f"{tagged[('Old', b)] - tagged[('Cohort', b)]} more clients {b}.")

deep = {}
for c in list(dict.fromkeys(DEEP_DIVE_CLIENTS + ["LOWER LLC"])):
    cells = inc[inc["Client"].eq(c)]
    slower = cells["Client_x_Cohort_ActualCPR"] < cells["Cohort_x_Cohort_ActualCPR"]
    lt = cells["Rule_bucket"].str.extract(r"LoanType=(\w+)")[0]
    lt_share = cells.groupby(lt)["Client_BeginPrinBal"].sum() / cells["Client_BeginPrinBal"].sum()
    deep[c] = {
        "rank": int(cohort_clients.loc[c, "VolumeRank"]),
        "funded": cohort_clients.loc[c, "FundedLoanAmt"],
        "cohorts": len(cells),
        "loan_months": int(cells["Client_LoanCt"].sum()),
        "actual": cohort_clients.loc[c, "Client_ActualCPR"],
        "cohort_bench": cohort_clients.loc[c, "Cohort_ActualCPR"],
        "cohort_cpr": cohort_clients.loc[c, "CPRScore"],
        "old_bench": old_clients.loc[c, "ModelCPR"],
        "old_cpr": old_clients.loc[c, "CPRScore"],
        "recap": cohort_clients.loc[c, "RecapScore"], "recap_old": old_clients.loc[c, "RecapScore"],
        "new": scores.loc[c, "new"], "old": scores.loc[c, "old"],
        "new_rel": scores.loc[c, "new_rel"], "old_rel": scores.loc[c, "old_rel"],
        "new_band": scores.loc[c, "new_band"], "old_band": scores.loc[c, "old_band"],
        "slower_n": int(slower.sum()),
        "slower_bal": cells.loc[slower, "Client_BeginPrinBal"].sum() / cells["Client_BeginPrinBal"].sum(),
        "main_type": lt_share.idxmax(), "main_type_share": lt_share.max(),
    }

PY, PH, PAD = 1.8, 4.15, 0.2
TAG_COL = {"Favourable": (GREEN, GREEN_TINT), "Neutral": (MUTE, "E9EBF2"), "Unfavourable": (RED, RED_TINT)}
TAG_LETTER = {"Favourable": "F", "Neutral": "N", "Unfavourable": "U"}
NAME_FIX = {"PRIMELENDING A PLAINSCAPITAL COMPANY": "PrimeLending", "LOANDEPOTCOM LLC": "loanDepot",
            "EVERETT FINANCIAL INC DBA SUPREME LENDING": "Supreme Lending", "DHI": "DHI Mortgage",
            "AVERRA FINANCIAL INC (SS 196043)": "Averra Financial (SS)", "USAA": "USAA", "NMSI INC": "NMSI",
            "NFM LENDING LLC": "NFM Lending", "LOWER LLC": "Lower LLC", "NVR MORTGAGE FINANCE INC": "NVR Mortgage",
            "PARAMOUNT RESIDENTIAL MORTGAGE GROUP INC": "Paramount Residential",
            "AMERICAN PACIFIC MORTGAGE CORPORATION": "American Pacific Mortgage"}


def short_name(s):
    if s in NAME_FIX:
        return NAME_FIX[s]
    s = re.sub(r"\s*\(.*?\)", "", s).split(" DBA ")[0]
    words = [w for w in s.split() if w not in {"LLC", "INC", "CORPORATION", "CORP", "COMPANY", "LP"}]
    caps = {"AFC", "GFS", "SSLP", "TPO", "LCNB", "USAA", "NMSI", "DHI", "NVR", "NFM", "GMFS", "VIP", "FCU", "TJC"}
    out = " ".join(w if w in caps else w.title() for w in words)
    return out.replace(" Of ", " of ").replace(" And ", " and ")


def set_alpha(shp, a):
    clr = shp.fill._xPr.find(qn("a:solidFill"))[0]
    clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(a * 100000))}))


def dashed(slide, x1, y1, x2, y2, color, width=1.0):
    ln = line(slide, x1, y1, x2, y2, color, width)
    ln.line.dash_style = MSO_LINE.DASH
    return ln


def tag_chip(slide, x, y, w, h, band, size=7.5):
    fc, tint = TAG_COL[band]
    box(slide, x, y, w, h, fill=tint, radius=min(w, h) / 2.2)
    text(slide, x, y, w, h, TAG_LETTER[band], size=size, bold=True, color=fc, align="center", valign="middle")


# ---------- client deep dives: cohort-by-cohort CPRScore build-up (cprAuditLong) ----------
DD_CLIENTS = ["AMERICAN PACIFIC MORTGAGE CORPORATION", "TOMO MORTGAGE LLC"]
_cpr = lambda end, beg: 1 - (end / beg) ** 12          # dictionary formula: annualised pooled run-off
_lo_spread = lambda rule: float(re.search(r"RateSpread \(([^,]+),", rule).group(1).replace("%", "").replace("Inf", "inf"))
_book = inc.assign(above=inc["Rule_bucket"].map(lambda r: _lo_spread(r) >= 0),
                   jumbo=inc["Rule_bucket"].str.contains(r"PrinBal >400k|PrinBal \(250k, \+Inf\]"))
BOOK_ABOVE = _book.loc[_book["above"], "Client_BeginPrinBal"].sum() / _book["Client_BeginPrinBal"].sum()
BOOK_JUMBO = _book.loc[_book["jumbo"], "Client_BeginPrinBal"].sum() / _book["Client_BeginPrinBal"].sum()


def cohort_label(rule):
    """'LoanType=CONV | FICOOrig >750 | OrigLoanTerm >240mo | PrinBal >400k | RateSpread (-0.5%, -0.25%]'
    -> 'Conv · FICO >750 · >$400K · −0.5 to −0.25%'."""
    out = []
    for part in [p.strip() for p in rule.split("|")]:
        k, _, v = part.partition("=") if "=" in part else part.partition(" ")
        if k == "LoanType":
            out.append({"CONV": "Conv"}.get(v, v))
        elif k == "FICOOrig":
            out.append(v.replace("<=", "≤"))
        elif k == "OrigLoanTerm" and "<=" in v:
            out.append("≤20yr")
        elif k == "PrinBal":
            v = v.replace("k", "K")
            out.append("any size" if v == "(-Inf, +Inf]" else
                       "≤$" + v.split(", ")[1].rstrip("]") if v.startswith("(-Inf") else
                       ">$" + v[1:].split(",")[0] if v.endswith("+Inf]") else
                       v.replace(">", ">$").replace("<=", "≤$") if v[0] in "<>" else "$" + v.replace("-", "–"))
        elif k == "RateSpread":
            lo, hi = [s.replace("%", "") for s in v.strip("(]").split(", ")]
            f = lambda s: {"-Inf": "below", "+Inf": "above"}.get(s, s.replace("-", "−"))
            out.append("any rate" if (lo, hi) == ("-Inf", "+Inf") else
                       f"≤ {f(hi)}%" if lo == "-Inf" else f"> {f(lo)}%" if hi == "+Inf" else f"{f(lo)} to {f(hi)}%")
    return " · ".join(out)


def loan_group(rule):
    if not rule.startswith("LoanType=CONV"):
        return "FHA / VA"
    return "Conv > $400K" if "PrinBal >400k" in rule else "Conv ≤ $400K"


def loan_groups(i):
    """Client CPR vs benchmark CPR (all lenders, same cohorts) within each loan group, plus its extra run-off."""
    g = i.groupby(i["Rule_bucket"].map(loan_group)).agg(
        cb=("Client_BeginPrinBal", "sum"), ce=("Client_EndPrinBal", "sum"), kb=("Cohort_BeginPrinBal", "sum"),
        ke=("Cohort_EndPrinBal", "sum"), extra=("extra", "sum"))
    g = g.reindex([k for k in ["Conv ≤ $400K", "Conv > $400K", "FHA / VA"] if k in g.index])
    g["share"] = g["cb"] / g["cb"].sum()
    g["client"], g["bench"] = _cpr(g["ce"], g["cb"]), _cpr(g["ke"], g["kb"])
    return g


def client_dive(c):
    rows = audit[audit["Client"].eq(c)]
    i = rows[rows["Included_Flag"].eq(1)].copy()
    cb, ce = i["Client_BeginPrinBal"].sum(), i["Client_EndPrinBal"].sum()
    kb, ke = i["Cohort_BeginPrinBal"].sum(), i["Cohort_EndPrinBal"].sum()
    i["peer_end"] = i["Client_BeginPrinBal"] * i["Cohort_EndPrinBal"] / i["Cohort_BeginPrinBal"]
    i["share"] = i["Client_BeginPrinBal"] / cb
    i["extra"] = i["peer_end"] - i["Client_EndPrinBal"]          # run-off beyond what peer speed implies
    client, bench = _cpr(ce, cb), _cpr(ke, kb)
    cc, oc = cohort_clients.loc[c], old_clients.loc[c]
    # tie-out: the audit rows reproduce the stored client CPR, benchmark and CPRScore
    assert abs(client - cc["Client_ActualCPR"]) < 1e-6 and abs(bench - cc["Cohort_ActualCPR"]) < 1e-6
    assert abs(client / bench - cc["CPRScore"]) < 1e-6
    top = set(i.nlargest(6, "share").index) | set(i["extra"].abs().nlargest(3).index)
    shown = i.loc[sorted(top, key=lambda k: -i.loc[k, "share"])]
    rest = i.drop(shown.index)
    return {
        "name": c, "rank": int(cc["VolumeRank"]), "funded": cc["FundedLoanAmt"], "cohorts": len(i),
        "loan_months": int(i["Client_LoanCt"].sum()), "excluded": rows.loc[rows["Included_Flag"].eq(0), "Client_BeginPrinBal"].sum()
        / rows["Client_BeginPrinBal"].sum(),
        "cb": cb, "ce": ce, "kb": kb, "ke": ke, "client": client, "bench": bench, "cpr_score": cc["CPRScore"],
        "groups": loan_groups(i),
        "paid": cb - ce, "paid_peer": cb - i["peer_end"].sum(),
        "faster_share": i.loc[i["Client_x_Cohort_ActualCPR"] > i["Cohort_x_Cohort_ActualCPR"], "share"].sum(),
        "above": i.loc[i["Rule_bucket"].map(lambda r: _lo_spread(r) >= 0), "share"].sum(),
        "jumbo": i.loc[i["Rule_bucket"].str.contains(r"PrinBal >400k|PrinBal \(250k, \+Inf\]"), "share"].sum(),
        "shown": shown, "rest": rest,
        "rest_client": _cpr(rest["Client_EndPrinBal"].sum(), rest["Client_BeginPrinBal"].sum()),
        "rest_peer": _cpr(rest["peer_end"].sum(), rest["Client_BeginPrinBal"].sum()),
        "recap": cc["RecapScore"], "recap_act": cc["ActualRecap"], "recap_bench": cc["CohortRecap"],
        "recap_n": int(cc["RecapLoanct"]),
        "old": {"bench": oc["ModelCPR"], "cpr": oc["CPRScore"], "recap": oc["RecapScore"], "recap_act": oc["ActualRecap"],
                "recap_bench": oc["CohortRecap"], "ts": oc["TotalScore"]},
        "new_ts": cc["TotalScore"], "new_rel": scores.loc[c, "new_rel"], "old_rel": scores.loc[c, "old_rel"],
        "new_band": scores.loc[c, "new_band"], "old_band": scores.loc[c, "old_band"],
    }


DIVES = [client_dive(c) for c in DD_CLIENTS]
BOOK_CPR = {"Old": float(old_total["CPRScore"]), "Cohort": float(cohort_total["CPRScore"])}
BOOK_RECAP = {"Old": float(old_total["RecapScore"]), "Cohort": float(cohort_total["RecapScore"])}


def deep_dive_slides(prs, page0):
    """One slide per deep-dive client: cohort table, the calculation, and why it landed where it did."""
    for n, d in enumerate(DIVES):
        s = new_slide(prs)
        tag(s, "CLIENT DEEP DIVE", INDIGO)
        nm = short_name(d["name"])
        faster = d["cpr_score"] - 1
        # what moved the tag between approaches: relative prepayment standing, or recapture
        o_rel_cpr, c_rel_cpr = d["old"]["cpr"] / BOOK_CPR["Old"], d["cpr_score"] / BOOK_CPR["Cohort"]
        o_rel_rec, c_rel_rec = d["old"]["recap"] / BOOK_RECAP["Old"], d["recap"] / BOOK_RECAP["Cohort"]
        recap_driven = abs(np.log(o_rel_rec / c_rel_rec)) > abs(np.log(c_rel_cpr / o_rel_cpr))
        if recap_driven:
            ttl = (f"{nm} tags {d['new_band']}: its loans prepay {pct(faster)} faster than look-alike loans, "
                   f"and its recapture edge narrows")
        else:
            ttl = (f"{nm} tags {d['new_band']}: its loans prepay {pct(faster)} faster than look-alike loans "
                   f"and recapture is below par")
        slide_title(s, ttl,
                    f"#{d['rank']} client by volume · {bn(d['funded'], 2)} funded · {d['loan_months']:,} loan-months "
                    f"in {d['cohorts']} cohorts · Funding {FUND_WINDOW}")

        # ----- Panel 1: cohort by cohort -----
        P1X, P1W = 0.5, 4.75
        box(s, P1X, PY, P1W, PH, fill=GRAYFILL, radius=0.08)
        ix, iw = P1X + PAD, P1W - 2 * PAD
        panel_header(s, ix, PY + 0.16, iw, 1, "COHORT BY COHORT",
                     f"Faster than peers in cohorts holding {pct(d['faster_share'])} of its balance")
        cols = [("Cohort", 0.4, "left"), ("Type · FICO · size · rate vs market", 1.98, "left"), ("Share", 0.4, "right"),
                ("Client", 0.48, "right"), ("Peers", 0.47, "right"), ("run-off", 0.62, "right")]
        xs = [ix + sum(w_ for _, w_, _ in cols[:k]) for k in range(len(cols))]
        hy = PY + 0.98
        text(s, xs[3], hy, cols[3][1] + cols[4][1], 0.16, "CPR IN COHORT", size=7, bold=True, color=INDIGO,
             align="right", char_spacing=1)
        text(s, xs[5], hy, cols[5][1], 0.16, "EXTRA", size=7, bold=True, color=INDIGO, align="right",
             char_spacing=1)
        for (lab, w_, al), x_ in zip(cols, xs):
            text(s, x_, hy + 0.16, w_, 0.17, lab, size=7.5, bold=True, color=MUTE, align=al)
        rh = 0.235
        body = [(r["Cohort_ID"], cohort_label(r["Rule_bucket"]), r["share"], r["Client_x_Cohort_ActualCPR"],
                 r["Cohort_x_Cohort_ActualCPR"], r["extra"]) for _, r in d["shown"].iterrows()]
        body.append((f"+{len(d['rest'])}", "All other cohorts", d["rest"]["share"].sum(), d["rest_client"],
                     d["rest_peer"], d["rest"]["extra"].sum()))
        for k, (cid, lab, sh, ccpr, pcpr, ex_) in enumerate(body):
            yy = hy + 0.36 + k * rh
            if k % 2 == 0:
                box(s, ix - 0.05, yy, iw + 0.1, rh, fill=WHITE)
            other = k == len(body) - 1
            fc = RED if ccpr > pcpr else GREEN
            vals = [(cid, {"color": MUTE}), (lab, {"italic": other}), (pct(sh), {}),
                    (pct(ccpr, 1), {"bold": True, "color": fc}), (pct(pcpr, 1), {"color": INK}),
                    ("≈ $0" if abs(ex_) < 5e4 else f"{'+' if ex_ >= 0 else '−'}${abs(ex_) / 1e6:,.1f}M",
                     {"bold": True, "color": fc})]
            for ci, (txt, o) in enumerate(vals):
                text(s, xs[ci], yy, cols[ci][1], rh, txt, size=7.5 if ci == 1 else 8, color=o.get("color", INK), bold=o.get("bold", False),
                     italic=o.get("italic", False), align=cols[ci][2], valign="middle")
        ty = hy + 0.36 + len(body) * rh
        line(s, ix, ty + 0.02, ix + iw, ty + 0.02, MUTE, 0.75)
        # pooled rows: client CPR vs benchmark CPR over every cohort, and the total of the extra run-off column
        extra_tot = d["shown"]["extra"].sum() + d["rest"]["extra"].sum()
        tot_rows = [
            [("", {}), (f"All {d['cohorts']} cohorts: client CPR vs benchmark", {"bold": True}), ("", {}),
             (pct(d["client"], 1), {"bold": True, "color": RED if d["client"] > d["bench"] else GREEN}),
             (pct(d["bench"], 1), {"bold": True, "color": INDIGO}), ("", {})],
            [("", {}), ("Total extra run-off (sum of all cohorts)", {"bold": True}), ("", {}), ("", {}), ("", {}),
             (f"{'+' if extra_tot >= 0 else '−'}${abs(extra_tot) / 1e6:,.1f}M",
              {"bold": True, "color": RED if extra_tot >= 0 else GREEN})],
        ]
        for r_, row in enumerate(tot_rows):
            for ci, (txt, o) in enumerate(row):
                text(s, xs[ci], ty + 0.04 + r_ * rh, cols[ci][1] if ci != 1 else cols[1][1] + cols[2][1], rh, txt,
                     size=7.5 if ci == 1 else 8, color=o.get("color", INK), bold=o.get("bold", False),
                     italic=o.get("italic", False), align=cols[ci][2], valign="middle")
        ny = ty + 0.04 + 2 * rh + 0.06
        text(s, ix, ny, iw, PY + PH - 0.06 - ny, [[
            ("Peers = all lenders’ loans in that cohort. Extra run-off = balance the client’s loans paid down beyond "
             "what the peers’ speed in that cohort implies (red = faster than peers). Rate = note rate vs market at lock."
             + (f" Loans with no cohort ({pct(d['excluded'], 1)} of balance) are left out." if d["excluded"] else ""), {}),
        ]], size=7.5, color=MUTE, spacing=0.92)

        # ----- Panel 2: the calculation (emphasis) -----
        P2X, P2W = P1X + P1W + 0.2, 3.7
        box(s, P2X, PY, P2W, PH, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
        ix, iw = P2X + PAD, P2W - 2 * PAD
        panel_header(s, ix, PY + 0.16, iw, 2, "THE CALCULATION", "From cohort balances to the tag")
        steps = [
            ("Client CPR", f"its loans, pooled over {d['cohorts']} cohorts",
             [(f"1 − ({bn(d['ce'], 2)} ÷ {bn(d['cb'], 2)})¹² = ", {}), (pct(d["client"], 1), {"bold": True, "color": INK})]),
            ("Benchmark CPR", f"every lender’s loans, same {d['cohorts']} cohorts",
             [(f"1 − ({bn(d['ke'], 0)} ÷ {bn(d['kb'], 0)})¹² = ", {}), (pct(d["bench"], 1), {"bold": True, "color": INK})]),
            ("CPRScore", "client ÷ benchmark (lower is better)",
             [(f"{pct(d['client'], 1)} ÷ {pct(d['bench'], 1)} = ", {}),
              (f"{d['cpr_score']:.2f}", {"bold": True, "color": INDIGO})]),
            ("TotalScore", f"recapture {pct(d['recap_act'], 1)} vs {pct(d['recap_bench'], 1)} = {d['recap']:.2f}",
             [(f"{d['recap']:.2f} ÷ {d['cpr_score']:.2f} = ", {}), (f"{d['new_ts']:.2f}", {"bold": True, "color": INDIGO})]),
            ("Tag", f"TotalScore ÷ portfolio score {Z['Cohort']:.2f}",
             [(f"{d['new_rel']:.2f}× ", {"bold": True, "color": TAG_COL[d['new_band']][0]}),
              (f"{'<' if d['new_rel'] < UNF_T else '≥' if d['new_rel'] >= FAV_T else 'between'} "
               f"{UNF_T:g}× → ", {}), (d["new_band"], {"bold": True, "color": TAG_COL[d['new_band']][0]})]),
        ]
        sy0, sgap = PY + 0.98, 0.04
        sh_ = 0.39
        for k, (lab, desc, val) in enumerate(steps):
            yy = sy0 + k * (sh_ + sgap)
            box(s, ix, yy, iw, sh_, fill=WHITE, radius=0.05)
            box(s, ix + 0.08, yy + (sh_ - 0.24) / 2, 0.24, 0.24, fill=INDIGO, radius=0.12)
            text(s, ix + 0.08, yy + (sh_ - 0.24) / 2, 0.24, 0.24, str(k + 1), size=8.5, bold=True, color=WHITE,
                 align="center", valign="middle")
            text(s, ix + 0.4, yy + 0.02, iw - 0.46, sh_ - 0.04, [
                [(lab + "  ", {"bold": True, "color": INK, "size": 9}), (desc, {"color": MUTE, "size": 7.5})],
                [(t, {**o, "size": 10}) for t, o in val],
            ], valign="middle", spacing=0.9)
        # the same client under the Old approach
        ty = sy0 + 5 * (sh_ + sgap) + 0.02
        cw0, cv = iw - 2 * 0.8, 0.8
        for (lab, col, al), x_, w_ in zip([("Same client, Old approach", INK, "left"), ("Old", OLD_C, "center"),
                                           ("Cohort", INDIGO, "center")], [ix, ix + cw0, ix + cw0 + cv], [cw0, cv, cv]):
            text(s, x_, ty, w_, 0.17, lab, size=8, bold=True, color=col, align=al)
        trows = [("CPR benchmark", pct(d["old"]["bench"], 1), pct(d["bench"], 1)),
                 ("CPRScore", f"{d['old']['cpr']:.2f}", f"{d['cpr_score']:.2f}"),
                 ("Recapture score", f"{d['old']['recap']:.2f}", f"{d['recap']:.2f}"),
                 ("TotalScore ÷ portfolio → tag", (f"{d['old_rel']:.2f}×", d["old_band"]),
                  (f"{d['new_rel']:.2f}×", d["new_band"]))]
        for k, (lab, o_, c_) in enumerate(trows):
            yy = ty + 0.19 + k * 0.18
            if k % 2 == 0:
                box(s, ix - 0.05, yy, iw + 0.1, 0.18, fill=WHITE)
            text(s, ix, yy, cw0, 0.18, lab, size=7.5, color=INK, valign="middle")
            for x_, v, col in [(ix + cw0, o_, OLD_C), (ix + cw0 + cv, c_, INDIGO)]:
                if isinstance(v, tuple):
                    text(s, x_, yy, cv / 2 + 0.12, 0.18, v[0], size=7.5, bold=True, color=col, align="right",
                         valign="middle")
                    tag_chip(s, x_ + cv / 2 + 0.17, yy + 0.015, 0.24, 0.15, v[1], size=7)
                else:
                    text(s, x_, yy, cv, 0.18, v, size=7.5, bold=True, color=col, align="center", valign="middle")

        # ----- Panel 3: where the extra prepayment comes from, by loan group (client CPR vs benchmark CPR) -----
        P3X = P2X + P2W + 0.2
        P3W = W - 0.5 - P3X
        box(s, P3X, PY, P3W, PH, fill=GRAYFILL, radius=0.08)
        ix, iw = P3X + PAD, P3W - 2 * PAD
        grp = d["groups"]
        all_fast = bool((grp["client"] > grp["bench"]).all())
        top_g = grp["extra"].idxmax()
        top_sh = grp.loc[top_g, "extra"] / grp["extra"].sum()
        head = ("Faster than its benchmark in every loan group" if all_fast else
                f"The extra prepayment is almost all in {top_g}" if top_sh >= 0.8 else
                f"Most of the extra prepayment is in {top_g}")
        panel_header(s, ix, PY + 0.16, iw, 3, "WHY IT LANDED HERE", head)
        text(s, ix, PY + 0.98, iw, 0.2, [[("■ ", {"color": RED}), ("Client CPR   ", {}), ("■ ", {"color": INDIGO}),
                                          ("Benchmark CPR   ", {}), ("same definitions as steps 1–2", {"italic": True})]],
             size=7.5, color=MUTE)
        vmax = float(max(grp["client"].max(), grp["bench"].max()))
        lab_w, val_w = 1.0, 0.42
        bx0, bwm = ix + lab_w, iw - lab_w - val_w - 0.02
        gy0, gh = PY + 1.2, 0.58
        for k, (g, r) in enumerate(grp.iterrows()):
            yy = gy0 + k * gh
            text(s, ix, yy, lab_w - 0.06, gh - 0.08, [
                [(g, {"bold": True, "color": INK, "size": 8.5})],
                [(f"{pct(r['share'])} of balance", {"color": MUTE, "size": 7.5})],
                [(f"{'+' if r['extra'] >= 0 else '−'}${abs(r['extra']) / 1e6:,.1f}M extra",
                  {"bold": True, "size": 7.5, "color": RED if r["extra"] >= 0 else GREEN})],
            ], valign="middle", spacing=0.88)
            for j, (v, col) in enumerate([(r["client"], RED if r["client"] > r["bench"] else GREEN),
                                          (r["bench"], INDIGO)]):
                by_ = yy + 0.08 + j * 0.21
                bw_ = max(bwm * v / vmax, 0.02)
                box(s, bx0, by_, bw_, 0.17, fill=col)
                text(s, bx0 + bw_ + 0.04, by_ - 0.02, val_w, 0.21, pct(v, 1), size=8, bold=True, color=col,
                     valign="middle")
        ty_ = gy0 + len(grp) * gh + 0.02
        line(s, ix, ty_, ix + iw, ty_, MUTE, 0.75)
        text(s, ix, ty_ + 0.04, iw, 0.2, [[("All loans: ", {"bold": True}),
                                          (pct(d["client"], 1), {"bold": True, "color": RED}), (" vs ", {}),
                                          (pct(d["bench"], 1), {"bold": True, "color": INDIGO}),
                                          (" → CPRScore ", {}), (f"{d['cpr_score']:.2f}", {"bold": True, "color": INDIGO})]],
             size=8.5, color=INK, valign="middle")
        r_top = grp.loc[top_g]
        if all_fast:
            read = [(f"Broad-based: ", {"bold": True}),
                    (f"{top_g} is {pct(r_top['share'])} of its balance but the largest source of extra run-off "
                     f"({r_top['client'] * 100:.1f}% vs {r_top['bench'] * 100:.1f}%)."
                     + (f" {pct(d['above'])} of its balance is above market rate (portfolio {pct(BOOK_ABOVE)}), "
                        "where loans refinance first." if d["above"] > BOOK_ABOVE + 0.05 else ""), {})]
        else:
            slow = grp[grp["client"] <= grp["bench"]].index.tolist()
            read = [(f"Concentrated: ", {"bold": True}),
                    (f"{pct(r_top['share'])} of its balance is {top_g}, prepaying at {r_top['client'] * 100:.1f}% vs "
                     f"{r_top['bench'] * 100:.1f}% for look-alike loans"
                     + (f"; its {' and '.join(slow)} loans are slower than benchmark." if slow else "."), {})]
        ry = ty_ + 0.28
        text(s, ix, ry, iw, 0.46, [read], size=7.5, color=INK, spacing=0.92)
        wy = ry + 0.48
        if recap_driven:
            why = [("Vs Old, the tag moves on recapture: ", {"bold": True}),
                   (f"its {pct(d['recap_act'], 0)} recapture now meets a {pct(d['recap_bench'], 1)} benchmark, not "
                    f"{pct(d['old']['recap_bench'], 1)}, so the score falls {d['old']['recap']:.2f} → {d['recap']:.2f} "
                    f"({d['recap_n']} payoffs).", {})]
        else:
            why = [("Vs Old, the tag moves on prepayment: ", {"bold": True}),
                   (f"the model benchmark ({pct(d['old']['bench'], 1)}) is low for most clients, so it sat mid-pack; "
                    f"against look-alike loans its speed stands out. Recapture is steady.", {})]
        text(s, ix, wy, iw, PY + PH - 0.05 - wy, [why], size=7.5, color=INK, spacing=0.92)
        impact_band(s, "THE RESULT", [
            (f"{nm}: {d['old_band']} under Old → {d['new_band']} under Cohort. ",
             {"bold": True, "color": WHITE}),
            (f"Client CPR {pct(d['client'], 1)} vs peers {pct(d['bench'], 1)}; recapture score {d['recap']:.2f}.",
             {"color": GOLD}),
        ])
        footer(s, page0 + n)
    return len(DIVES)


def results_slides(prs, page0):
    """Slides: score impact; why scores move; client deep dives; stability. Returns slide count."""

    # =====================================================================
    # SCORE IMPACT — cross-tab, every client, top 25 lenders
    # =====================================================================
    s = new_slide(prs)
    tag(s, "SCORE IMPACT", INDIGO)
    slide_title(
        s,
        f"The two approaches agree on {pct(same / n_clients)} of clients; Cohort tags {down} clients lower "
        f"and {up} higher",
        f"{n_clients} clients · Funding {FUND_WINDOW} · Tags vs own portfolio: "
        f"Favourable ≥ {FAV_T:g}×; Unfavourable < {UNF_T:g}×",
    )

    # ----- Panel 1: head-to-head (rows Old, columns Cohort) -----
    P1X, P1W = 0.5, 3.3
    box(s, P1X, PY, P1W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P1X + PAD, P1W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 1, "HEAD-TO-HEAD", f"{pct(same / n_clients)} keep the same tag")
    text(s, ix, PY + 0.8, iw, 0.2, [[("Rows: ", {}), ("Old", {"bold": True, "color": OLD_C}),
         (" tag   ·   Columns: ", {}), ("Cohort", {"bold": True, "color": INDIGO}), (" tag", {})]],
         size=9, color=MUTE)
    gy = PY + 1.04
    lab_w, rh, hh = 0.62, 0.42, 0.24
    cw = (iw - lab_w) / 3
    short = {"Favourable": "Fav.", "Neutral": "Neutral", "Unfavourable": "Unfav."}
    for j, b in enumerate(BANDS):
        text(s, ix + lab_w + j * cw, gy, cw, hh, short[b], size=9.5, bold=True, color=INDIGO,
             align="center", valign="middle")
    for i, rb in enumerate(BANDS):
        ry = gy + hh + i * rh
        text(s, ix, ry, lab_w, rh, short[rb], size=9.5, bold=True, color=OLD_C, valign="middle")
        for j, cb in enumerate(BANDS):
            v = int(xtab.loc[rb, cb])
            if i == j:
                fill, fc = INDIGO, WHITE
            elif abs(i - j) == 1:
                fill, fc = (AMBER_TINT, AMBER) if v >= 50 else (WHITE, MUTE)
            else:
                fill, fc = (RED_TINT, RED) if v else (WHITE, MUTE)
            hl = rb == "Unfavourable" and cb == "Favourable"
            box(s, ix + lab_w + j * cw + 0.03, ry + 0.03, cw - 0.06, rh - 0.06, fill=fill,
                line=INDIGO if hl else (LINEGRAY if fill == WHITE else None),
                line_w=2.25 if hl else 0.75, radius=0.05)
            text(s, ix + lab_w + j * cw, ry, cw, rh, f"{v:,}", size=13, bold=True, color=fc,
                 align="center", valign="middle")
    ly = gy + hh + 3 * rh + 0.08
    text(s, ix, ly, iw, 0.22, [[
        ("■ ", {"color": INDIGO}), (f"Same {pct(same / n_clients)}", {"bold": True}),
        ("   ■ ", {"color": AMBER}), (f"One step {pct(one_step / n_clients)}", {}),
        ("   ■ ", {"color": RED}), (f"Opposite {pct(opposite / n_clients, 1)}", {}),
    ]], size=9.5, color=INK)
    notes = [[(f"Cohort tags {unf_total['Cohort']} clients Unfavourable vs {unf_total['Old']}, "
               f"and {fav_total['Cohort']} Favourable vs {fav_total['Old']}.", {"bold": True})]]
    if zero_unf:
        notes.append([(f"{zero_unf} of Cohort’s Unfavourable clients had no prepayments yet", {"bold": True, "space_before": 4}),
                      (f" and get a default score of {zero_default:.2f}; this includes {rev_zero['FU']} of the "
                       f"{fav_to_unf} Fav. → Unfav. reversals.", {})])
    notes.append([(f"{unf_to_fav} move Unfav. → Fav.", {"bold": True, "space_before": 4}),
                  (f" (outlined), led by {short_name(flip_leader)} (#{int(scores.loc[flip_leader, 'rank'])}); "
                   "next slide shows why. All reversals: appendix.", {})])
    text(s, ix, ly + 0.3, iw, 5.88 - ly - 0.3, notes, size=9, color=INK, spacing=0.96)

    # ----- Panel 2: every client, relative TotalScore Old vs Cohort (emphasis) -----
    P2X, P2W = P1X + P1W + 0.2, 4.95
    box(s, P2X, PY, P2W, PH, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = P2X + PAD, P2W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 2, "EVERY CLIENT",
                 f"{pct(share_higher)} of clients stand higher under Cohort")
    px0, px1 = ix + 0.42, ix + iw - 0.08
    py0, py1 = PY + 1.02, 5.22
    LO, HI = 0.1, 10.0
    fx = lambda v: px0 + (px1 - px0) * (np.log10(min(max(v, LO), HI)) - np.log10(LO)) / 2
    fy = lambda v: py1 - (py1 - py0) * (np.log10(min(max(v, LO), HI)) - np.log10(LO)) / 2
    box(s, px0, py0, px1 - px0, py1 - py0, fill=WHITE)
    box(s, px0, py0, px1 - px0, fy(FAV_T) - py0, fill="F1F8F4")
    box(s, px0, fy(UNF_T), px1 - px0, py1 - fy(UNF_T), fill="FCF1F3")
    for t in (FAV_T, UNF_T):
        dashed(s, px0, fy(t), px1, fy(t), REFGRAY, 0.75)
        dashed(s, fx(t), py0, fx(t), py1, REFGRAY, 0.75)
    dashed(s, fx(LO), fy(LO), fx(HI), fy(HI), MUTE, 1.0)
    for v in (0.1, UNF_T, FAV_T, 10):
        text(s, px0 - 0.42, fy(v) - 0.08, 0.36, 0.16, f"{v:g}×", size=7.5, color=INDIGO, align="right")
        text(s, fx(v) - 0.25, py1 + 0.03, 0.5, 0.16, f"{v:g}×", size=7.5, color=OLD_C, align="center")
    text(s, px0 + 0.06, py0 + 0.04, 1.4, 0.16, "Favourable under Cohort", size=7.5, bold=True, color=GREEN)
    text(s, px1 - 1.66, py1 - 0.2, 1.6, 0.16, "Unfavourable under Cohort", size=7.5, bold=True, color=RED,
         align="right")
    fmax = scores["funded"].max()
    for _, r in scores.sort_values("funded").iterrows():
        d = 0.04 + 0.2 * np.sqrt(r["funded"] / fmax)
        col = {1: INDIGO, 0: "8D95AE", -1: AMBER}[int(r["move"])]
        dot = box(s, fx(r["old_rel"]) - d / 2, fy(r["new_rel"]) - d / 2, d, d, fill=col, line=WHITE, line_w=0.25,
                  shape=MSO_SHAPE.OVAL)
        set_alpha(dot, 0.6 if d < 0.1 else 0.85)
    labels = [("ZILLOW HOME LOANS LLC", "Zillow", 0.2, 0.2, "left"), ("LENNAR MORTGAGE LLC", "Lennar", 0.16, -0.16, "left"),
              (flip_leader, short_name(flip_leader), -1.12, -0.12, "right"),
              (down_leader, short_name(down_leader).split()[0] + " " + short_name(down_leader).split()[1]
               if len(short_name(down_leader).split()) > 1 else short_name(down_leader), 0.16, 0.3, "left")]
    for c, lab, dx, dy, al in labels:
        r = scores.loc[c]
        text(s, fx(r["old_rel"]) + dx, fy(r["new_rel"]) + dy - 0.07, 1.0, 0.16,
             [[(lab, {"bold": True}), (f" #{int(r['rank'])}", {"color": MUTE})]], size=8, color=INK, align=al)
    text(s, px0, py1 + 0.19, px1 - px0, 0.16, "Old TotalScore ÷ Old portfolio score  →", size=8, bold=True,
         color=OLD_C, align="right")
    text(s, ix, py0 - 0.22, 3.2, 0.16, "↑  Cohort TotalScore ÷ Cohort portfolio score", size=8, bold=True,
         color=INDIGO)
    clamped = int(((scores[["new_rel", "old_rel"]] < LO) | (scores[["new_rel", "old_rel"]] > HI)).any(axis=1).sum())
    text(s, ix, 5.62, iw, 0.2, [[
        ("● ", {"color": INDIGO}), ("Better tag   ", {}), ("● ", {"color": "8D95AE"}), ("Same tag   ", {}),
        ("● ", {"color": AMBER}), ("Lower tag   ", {}),
        (f"Size = volume · diagonal = same standing · {clamped} off-scale", {"color": MUTE}),
    ]], size=8, color=INK)

    # ----- Panel 3: top 25 lenders, tags only -----
    P3X = P2X + P2W + 0.2
    P3W = W - 0.5 - P3X
    box(s, P3X, PY, P3W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P3X + PAD, P3W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 3, "TOP 25 LENDERS", f"{pct(top25_share)} of funded volume")
    cols = [("#", 0.26, "left"), ("Lender", 1.72, "left"), ("OLD", 0.55, "center"), ("COHORT", 0.6, "center"),
            ("", iw - 0.26 - 1.72 - 1.15 - 0.12, "center")]
    xs, cx_ = [], ix
    for _, w_, _ in cols:
        xs.append(cx_)
        cx_ += w_ + (0.04 if _ else 0)
    hy = PY + 0.78
    for (lab, w_, al), x_ in zip(cols, xs):
        text(s, x_, hy, w_, 0.18, lab, size=7.5, bold=True, align=al, char_spacing=0.5,
             color=OLD_C if lab == "OLD" else INDIGO if lab == "COHORT" else MUTE, valign="middle")
    rh = 0.116
    for k, (name, r) in enumerate(top25.iterrows()):
        yy = hy + 0.2 + k * rh
        if k % 2 == 0:
            box(s, ix - 0.05, yy, iw + 0.1, rh, fill=WHITE)
        text(s, xs[0], yy, cols[0][1], rh, str(int(r["rank"])), size=7.5, color=MUTE, valign="middle")
        text(s, xs[1], yy, cols[1][1], rh, short_name(name), size=7.5, color=INK, valign="middle",
             bold=name in DEEP_DIVE_CLIENTS)
        for ci, key in [(2, "old_band"), (3, "new_band")]:
            cw_ = 0.3
            tag_chip(s, xs[ci] + (cols[ci][1] - cw_) / 2, yy + 0.008, cw_, rh - 0.016, r[key], size=7)
        mv = int(r["move"])
        text(s, xs[4], yy, cols[4][1], rh, {1: "▲", 0: "–", -1: "▼"}[mv], size=7, bold=True,
             color={1: INDIGO, 0: MUTE, -1: AMBER}[mv], align="center", valign="middle")
    text(s, ix, hy + 0.2 + 25 * rh + 0.04, iw, 0.18, [[
        ("F ", {"color": GREEN, "bold": True}), ("Fav.  ", {}), ("N ", {"color": MUTE, "bold": True}), ("Neutral  ", {}),
        ("U ", {"color": RED, "bold": True}), ("Unfav.   ", {}),
        (f"▲ {top25_moves['up']} better · – {top25_moves['same']} same · ▼ {top25_moves['down']} lower",
         {"bold": True}),
    ]], size=7.5, color=MUTE)

    lead = ("Cohort is the stricter read. " if unf_total["Cohort"] > unf_total["Old"]
            else "Cohort is the more generous read. ")
    impact_band(s, "THE IMPACT", [
        (lead, {"bold": True, "color": WHITE}),
        (f"{unf_total['Cohort']} vs {unf_total['Old']}", {"bold": True, "color": GOLD}),
        (" clients tag Unfavourable" + (f" ({zero_unf} of them with no prepayments yet)" if zero_unf else "")
         + "; among the top 25 lenders, ", {}),
        (f"{top25_moves['up']} improve and {top25_moves['down']} slip", {"bold": True, "color": GOLD}),
        (f", and {opposite} clients reverse between Favourable and Unfavourable.", {}),
    ], size=11.5)
    footer(s, page0)

    # =====================================================================
    # WHY SCORES MOVE — the benchmark, and two clients through the calculation
    # =====================================================================
    s = new_slide(prs)
    tag(s, "WHY SCORES MOVE", INDIGO)
    slide_title(
        s,
        "Different prepayment benchmarks reorder clients; the gap varies by loan type",
        "Old uses model expectations; Cohort uses observed look-alike loan performance. Tags reflect relative standing.",
    )
    BAND4_Y, BAND4_H = 6.36, 0.54
    PH4 = BAND4_Y - 0.14 - PY
    if DECK_MODE["compact"]:
        STRETCH[s.slide_id] = (PY, PY + PH4, CONTENT_BOTTOM)

    # ----- Panel 1: benchmark by loan type, and the resulting tag mix -----
    P1X, P1W = 0.5, 3.9
    box(s, P1X, PY, P1W, PH4, fill=GRAYFILL, radius=0.08)
    ix, iw = P1X + PAD, P1W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 1, "THE BENCHMARK", "The benchmark gap differs by loan type")
    by = PY + 1.08
    text(s, ix, by, iw, 0.2, [[("CPR benchmark by loan type   ", {"bold": True, "color": INK, "size": 9.5}),
                               ("■ ", {"color": OLD_C}), ("Old  ", {}), ("■ ", {"color": INDIGO}), ("Cohort", {})]],
         size=8.5, color=MUTE)
    lw_ = 0.95
    bmax = iw - lw_ - 0.5
    vmax = max(max(o_, c_) for _, o_, c_ in bench_lt)
    for k, (lab, o_, c_) in enumerate(bench_lt):
        yy = by + 0.26 + k * 0.36
        text(s, ix, yy, lw_ - 0.05, 0.3, lab, size=9, bold=True, color=INK, valign="middle")
        for j, (v, col) in enumerate([(o_, OLD_C), (c_, INDIGO)]):
            bw = bmax * v / vmax
            box(s, ix + lw_, yy + 0.02 + j * 0.14, bw, 0.12, fill=col)
            text(s, ix + lw_ + bw + 0.05, yy - 0.01 + j * 0.14, 0.5, 0.18, pct(v, 1), size=8, bold=True, color=col,
                 valign="middle")
    exc = "; ".join(f"for {lab} it sits above ({pct(m, 1)} vs {pct(act, 1)})" for lab, m, act in bench_exceptions)
    text(s, ix, by + 1.36, iw, 0.36, [[
        (f"Old is lower for {bench_lower} of {bench_n} clients", {"bold": True}),
        (f" ({pct(bench_lower_vol, 1)} of volume)" + (f"; {exc}." if exc else "."), {}),
    ]], size=8.5, color=INK, spacing=0.95)

    dy0 = by + 1.8
    text(s, ix, dy0, iw, 0.2, [[
        ("Clients by tag   ", {"bold": True, "color": INK, "size": 9.5}),
        ("■ ", {"color": GREEN}), ("Fav.  ", {}), ("■ ", {"color": REFGRAY}), ("Neutral  ", {}),
        ("■ ", {"color": RED}), ("Unfav.", {}),
    ]], size=8.5, color=MUTE)
    lw2 = 0.62
    bx0, bw_ = ix + lw2, iw - lw2
    for k, (a, col) in enumerate([("Old", OLD_C), ("Cohort", INDIGO)]):
        ry = dy0 + 0.26 + k * 0.36
        text(s, ix, ry, lw2 - 0.06, 0.28, a, size=9.5, bold=True, color=col, valign="middle")
        sx = bx0
        for b, fill, fc in [("Favourable", GREEN, WHITE), ("Neutral", REFGRAY, INK), ("Unfavourable", RED, WHITE)]:
            sw = bw_ * tag_share[a][b]
            if sw <= 0:
                continue
            box(s, sx, ry, sw, 0.28, fill=fill)
            if sw > 0.32:
                text(s, sx, ry, sw, 0.28, pct(tag_share[a][b]), size=9, bold=True, color=fc, align="center",
                     valign="middle")
            sx += sw
    text(s, ix, dy0 + 1.02, iw, PY + PH4 - 0.08 - dy0 - 1.02, [[
        ("Each approach is tagged against its own portfolio score, ", {"bold": True}),
        ("so what changes a tag is a client’s standing relative to the rest of the book.", {}),
    ]], size=8.5, color=INK, spacing=0.95)

    # ----- Panel 2: two clients, benchmark → CPR score → TotalScore → tag (emphasis) -----
    P2X = P1X + P1W + 0.2
    P2W = W - 0.5 - P2X
    box(s, P2X, PY, P2W, PH4, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = P2X + PAD, P2W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 2, "INSIDE THE CALCULATION",
                 "How the benchmark flows through to the tag: two clients")
    text(s, ix, PY + 0.8, iw, 0.2, [[
        ("Benchmark  ", {"bold": True, "color": INK}), ("Old", {"bold": True, "color": OLD_C}),
        (" = pre-stored model CPR for the client’s loans   ·   ", {}), ("Cohort", {"bold": True, "color": INDIGO}),
        (" = actual CPR of look-alike loans in the client’s cohorts", {}),
    ]], size=9, color=MUTE)

    cgap = 0.12
    cy_top = PY + 1.08
    ch_ = (PY + PH4 - 0.16 - cy_top - cgap) / 2
    LZ = 2.7
    cols = [("", 0.66), ("Client CPR", 0.56), ("÷ Benchmark", 0.68), ("= CPR score", 0.6),
            ("Recap ÷ CPR score", 0.9), ("= TotalScore", 0.62), ("÷ Portfolio = tag", 0.0)]
    for k, c in enumerate(DEEP_DIVE_CLIENTS):
        d = deep[c]
        assert abs(old_clients.loc[c, "ActualCPR"] - d["actual"]) < 5e-4
        cx, cy = ix, cy_top + k * (ch_ + cgap)
        box(s, cx, cy, iw, ch_, fill=WHITE, radius=0.08)
        meta = f"#{d['rank']} · ${d['funded'] / 1e6:,.0f}M"
        if d["main_type_share"] >= 0.9:
            meta += f" · {pct(d['main_type_share'])} {'Conv' if d['main_type'] == 'CONV' else d['main_type']}"
        text(s, cx + 0.16, cy + 0.06, LZ - 0.2, 0.4, [
            [(short_name(c), {"bold": True, "size": 12, "font": FH})], [(meta, {"size": 8, "color": MUTE})],
        ], valign="top", spacing=0.9)
        ax_max = 0.35 if max(d["actual"], d["cohort_bench"], d["old_bench"]) > 0.15 else 0.2
        step = 0.1
        ax0, ax1 = cx + 0.28, cx + LZ - 0.1
        ay = cy + ch_ * 0.65
        xpos = lambda v: ax0 + (ax1 - ax0) * v / ax_max
        line(s, ax0, ay, ax1, ay, REFGRAY, 1.5)
        for v in np.arange(0, ax_max + 1e-9, step):
            text(s, xpos(v) - 0.25, ay + 0.3, 0.5, 0.15, f"{v * 100:.0f}%", size=7.5, color=MUTE, align="center")
        close = abs(xpos(d["old_bench"]) - xpos(d["cohort_bench"])) < 0.75
        for v, col, lab in [(d["old_bench"], OLD_C, "Old"), (d["cohort_bench"], INDIGO, "Cohort")]:
            line(s, xpos(v), ay - 0.18, xpos(v), ay + 0.1, col, 2.25)
            shift = (0 if not close else (-0.3 if lab == "Old" else 0.3) * (1 if d["old_bench"] < d["cohort_bench"] else -1))
            text(s, xpos(v) - 0.5 + shift, ay - 0.52, 1.0, 0.33, [
                [(lab, {"size": 8, "color": col})], [(pct(v, 1), {"size": 10.5, "bold": True, "color": col})],
            ], align="center", valign="bottom", spacing=0.9)
        dd = 0.18
        box(s, xpos(d["actual"]) - dd / 2, ay - dd / 2, dd, dd, fill=INK, shape=MSO_SHAPE.OVAL)
        text(s, xpos(d["actual"]) - 0.75, ay + 0.1, 1.5, 0.18, [[
            ("Client ", {"size": 8, "color": INK}), (pct(d["actual"], 1), {"size": 9.5, "bold": True}),
        ]], align="center", valign="top")

        rx0 = cx + LZ + 0.12
        rw = iw - LZ - 0.12 - 0.12
        widths = [w_ for _, w_ in cols[:-1]] + [rw - sum(w_ for _, w_ in cols[:-1])]
        xs = [rx0 + sum(widths[:i]) for i in range(len(widths))]
        gy_ = cy + 0.1
        for (lab, span0, span1) in [("1 · CPR SCORE", 1, 3), ("2 · TOTALSCORE", 4, 5), ("3 · TAG", 6, 6)]:
            gx0, gx1 = xs[span0] + 0.04, xs[span1] + widths[span1] - 0.04
            text(s, gx0, gy_, gx1 - gx0, 0.16, lab, size=7.5, bold=True, color=INDIGO, align="center",
                 char_spacing=1)
            line(s, gx0, gy_ + 0.19, gx1, gy_ + 0.19, "BEB9F6", 1)
        for (lab, _), x_, w_ in zip(cols, xs, widths):
            text(s, x_, gy_ + 0.22, w_, 0.18, lab, size=7.5, color=MUTE, align="center")
        rh_ = (ch_ - 0.54 - 0.3 - 0.05) / 2
        for j, (lab, a, bv, cpr, rc, ts, rel, band, col, tint) in enumerate([
            ("OLD", "Old", d["old_bench"], d["old_cpr"], d["recap_old"], d["old"], d["old_rel"], d["old_band"],
             OLD_C, OLD_TINT),
            ("COHORT", "Cohort", d["cohort_bench"], d["cohort_cpr"], d["recap"], d["new"], d["new_rel"],
             d["new_band"], INDIGO, INDIGO_TINT),
        ]):
            ry = cy + 0.52 + j * (rh_ + 0.05)
            assert abs(rc / cpr - ts) < 5e-3 and abs(ts / Z[a] - rel) < 1e-9, (c, lab)
            box(s, rx0, ry, rw, rh_, fill=tint, radius=0.05)
            text(s, xs[0] + 0.06, ry, widths[0], rh_, [
                [(lab, {"size": 8, "bold": True, "color": col, "char_spacing": 1})],
                [(f"portfolio {Z[a]:.2f}", {"size": 7, "color": MUTE})],
            ], valign="middle", spacing=0.85)
            text(s, xs[1], ry, widths[1], rh_, pct(d["actual"], 1), size=10, color=INK, align="center",
                 valign="middle")
            box(s, xs[2] + 0.07, ry + 0.05, widths[2] - 0.14, rh_ - 0.1, fill=WHITE, line=col, line_w=1.25,
                radius=0.05)
            text(s, xs[2], ry, widths[2], rh_, pct(bv, 1), size=10.5, bold=True, color=col, align="center",
                 valign="middle")
            text(s, xs[3], ry, widths[3], rh_, f"{cpr:.2f}", size=10.5, bold=True, color=col, align="center",
                 valign="middle")
            text(s, xs[4], ry, widths[4], rh_, [[(f"{rc:.2f} ÷ ", {"color": INK}),
                                                  (f"{cpr:.2f}", {"bold": True, "color": col})]],
                 size=10, align="center", valign="middle")
            text(s, xs[5], ry, widths[5], rh_, f"{ts:.2f}", size=10.5, bold=True, color=col, align="center",
                 valign="middle")
            text(s, xs[6], ry, widths[6] - 0.04, rh_, [
                [(f"{rel:.2f}×", {"size": 11, "bold": True, "color": TAG_COL[band][0]})],
                [(band, {"size": 7.5, "bold": True, "color": TAG_COL[band][0]})],
            ], align="center", valign="middle", spacing=0.8)
        if d["main_type"] != "CONV" and d["main_type_share"] >= 0.9:
            ev = [(f"{d['main_type']} peers prepaid at {pct(d['cohort_bench'], 1)}, not {pct(d['old_bench'], 1)}",
                   {"bold": True}), (f": CPR score {d['old_cpr']:.2f} → {d['cohort_cpr']:.2f}", {})]
        elif d["cohort_cpr"] > 1:
            ev = [(f"Prepays {pct(d['cohort_cpr'] - 1)} faster than look-alike loans", {"bold": True}),
                  (f" ({pct(d['actual'], 1)} vs {pct(d['cohort_bench'], 1)})", {})]
        else:
            ev = [("Slower than peers in ", {}), (f"{d['slower_n']} of {d['cohorts']} cohorts", {"bold": True}),
                  (f" ({pct(d['slower_bal'])} of balance): CPR score below 1", {})]
        text(s, rx0, cy + ch_ - 0.28, rw, 0.22, [ev], size=8.5, color=INK, valign="middle")

    impact_band(s, "WHY THEY DIFFER", [
        ("Same client speed, a market-based benchmark. ", {"bold": True, "color": WHITE}),
        (f"The Old model expects similar speed across loan types while actual speed differs (e.g. {worst_lt[0]} "
         f"{pct(worst_lt[2], 1)} vs {pct(worst_lt[1], 1)} expected); Cohort compares each client with loans like its own.",
         {}),
    ], y=BAND4_Y, h=BAND4_H, size=10.5)
    footer(s, page0 + 1)
    n_dd = deep_dive_slides(prs, page0 + 2)
    page0 += n_dd

    # =====================================================================
    # STABILITY — do tags hold in the next period?
    # =====================================================================
    s = new_slide(prs)
    tag(s, "STABILITY", INDIGO)
    slide_title(
        s,
        f"Cohort tags hold over the next {N_MONTHS_WORD} months for {pct(tag_rate['Cohort'])} of tagged clients, "
        f"against {pct(tag_rate['Old'])} under the Old approach",
        f"Tags from {FB[2]} checked against the same clients’ {FC[2]} tags, each against its own portfolio score.",
    )
    box(s, 0.5, PY, W - 1.0, PH, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = 0.5 + PAD, W - 1.0 - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 1, "STABILITY", "Cohort tags are confirmed more often by what clients do next"
                 if tag_rate["Cohort"] >= tag_rate["Old"] else "Old tags are confirmed more often by what clients do next")

    LW = 5.0
    hy = PY + 0.86
    box(s, ix, hy, LW, 1.42, fill=WHITE, radius=0.07)
    text(s, ix + 0.2, hy + 0.12, LW - 0.4, 0.2, "TOTAL PREDICTION STABILITY", size=9.5, bold=True, color=MUTE,
         char_spacing=1.2)
    text(s, ix + 0.2, hy + 0.32, LW - 0.4, 0.62, [[
        (pct(tag_rate["Cohort"], 1), {"size": 38, "bold": True, "color": INDIGO}),
        ("   vs   ", {"size": 12, "color": MUTE}),
        (pct(tag_rate["Old"], 1), {"size": 38, "bold": True, "color": OLD_C}),
    ]], valign="middle")
    text(s, ix + 0.2, hy + 0.96, LW - 0.4, 0.4, [[
        (f"{total_hits['Cohort']} of {total_tagged['Cohort']} vs {total_hits['Old']} of {total_tagged['Old']}",
         {"bold": True, "color": INK}),
        (f" {F_BASE} Fav. / Unfav. tags confirmed in {F_COMP}", {}),
    ]], size=9.5, color=MUTE, spacing=0.95)

    ty, th, tg = hy + 1.42 + 0.14, 1.27, 0.12
    tw = (LW - 2 * tg) / 3
    tiles = [
        ("FAVOURABLE CONFIRMED", pct(hit_rate[("Cohort", "Favourable")], 1), pct(hit_rate[("Old", "Favourable")], 1),
         f"{hit[('Cohort', 'Favourable')]} of {tagged[('Cohort', 'Favourable')]} vs "
         f"{hit[('Old', 'Favourable')]} of {tagged[('Old', 'Favourable')]} tagged Favourable"),
        ("UNFAVOURABLE CONFIRMED", pct(hit_rate[("Cohort", "Unfavourable")], 1),
         pct(hit_rate[("Old", "Unfavourable")], 1),
         f"{hit[('Cohort', 'Unfavourable')]} of {tagged[('Cohort', 'Unfavourable')]} vs "
         f"{hit[('Old', 'Unfavourable')]} of {tagged[('Old', 'Unfavourable')]} tagged Unfavourable"),
        ("GOOD CLIENTS MISSED", pct(good_flagged_bad["Cohort"] / tagged[("Cohort", "Unfavourable")], 1),
         pct(good_flagged_bad["Old"] / tagged[("Old", "Unfavourable")], 1),
         f"{good_flagged_bad['Cohort']} vs {good_flagged_bad['Old']} Unfavourable tags turned Favourable "
         "(lower is better)"),
    ]
    for k, (lab, cv, ov, desc) in enumerate(tiles):
        tx = ix + k * (tw + tg)
        box(s, tx, ty, tw, th, fill=WHITE, radius=0.06)
        text(s, tx + 0.12, ty + 0.08, tw - 0.2, 0.32, lab, size=8, bold=True, color=MUTE, char_spacing=0.8,
             spacing=0.9)
        text(s, tx + 0.12, ty + 0.4, tw - 0.2, 0.34, [[
            (cv, {"size": 15, "bold": True, "color": INDIGO}), (" vs ", {"size": 8.5, "color": MUTE}),
            (ov, {"size": 15, "bold": True, "color": OLD_C}),
        ]], valign="middle")
        text(s, tx + 0.12, ty + 0.78, tw - 0.2, 0.44, desc, size=8.5, color=MUTE, spacing=0.92)

    RX = ix + LW + 0.3
    RW = ix + iw - RX
    text(s, RX, hy, RW, 0.22, [[
        (f"Where each {F_BASE} tag landed in {F_COMP}     ", {"bold": True, "color": INK, "size": 10.5}),
        ("■ ", {"color": GREEN}), ("Confirmed   ", {}), ("■ ", {"color": REFGRAY}), ("Neutral   ", {}),
        ("■ ", {"color": AMBER}), ("Reversed", {}),
    ]], size=9, color=MUTE)
    rows = [("Favourable", "Cohort"), ("Favourable", "Old"), ("Unfavourable", "Cohort"), ("Unfavourable", "Old")]
    lab_w = 1.9
    bar_x0, bar_w_max = RX + lab_w, RW - lab_w - 0.62
    max_n = max(tagged.values())
    row_y = hy + 0.4
    for r, (band, a) in enumerate(rows):
        yy = row_y + r * 0.46 + (0.16 if r >= 2 else 0)
        col = INDIGO if a == "Cohort" else OLD_C
        opp = "Unfavourable" if band == "Favourable" else "Favourable"
        n = tagged[(a, band)]
        text(s, RX, yy, lab_w - 0.1, 0.34, [[
            (f"{band} · ", {"color": INK}), (a, {"bold": True, "color": col}), (f"  {n}", {"color": MUTE}),
        ]], size=10, valign="middle")
        segs = [(hit[(a, band)], GREEN, WHITE), (int(next_tr[a].loc[band, "Neutral"]), REFGRAY, INK),
                (int(next_tr[a].loc[band, opp]), AMBER, WHITE)]
        sx = bar_x0
        for v, fill, fc in segs:
            sw = bar_w_max * v / max_n
            if sw <= 0:
                continue
            box(s, sx, yy + 0.03, sw, 0.28, fill=fill)
            if sw > 0.25:
                text(s, sx, yy + 0.03, sw, 0.28, str(v), size=9.5, bold=True, color=fc, align="center",
                     valign="middle")
            sx += sw
        text(s, sx + 0.08, yy, 0.6, 0.34, pct(hit_rate[(a, band)]), size=10.5, bold=True, color=GREEN,
             valign="middle")
    cap_y = row_y + 4 * 0.46 + 0.28
    text(s, RX, cap_y, RW, 5.85 - cap_y, [
        [("Bar length = clients given that tag; % = share of that tag confirmed. ", {"color": MUTE}),
         (caption_lead, {"bold": True})],
        *[[(t, {"space_before": 4})] for t in caption_notes],
    ], size=9.5, color=INK, spacing=0.98)

    band_runs = [("Cohort tags are the better read on what clients do next: " if tag_rate["Cohort"] >= tag_rate["Old"]
                  else "Old tags hold slightly better overall: ", {"bold": True, "color": WHITE}),
                 (f"{pct(tag_rate['Cohort'], 1)} vs {pct(tag_rate['Old'], 1)}", {"bold": True, "color": GOLD}),
                 (" of tags confirmed", {})]
    band_runs += [(". Cohort confirms Unfavourable at ", {}),
                  (_hr('Cohort', 'Unfavourable'), {"bold": True, "color": GOLD}),
                  ("; Old at ", {}), (_hr('Old', 'Unfavourable'), {"bold": True, "color": GOLD})]
    r_missed = good_flagged_bad["Cohort"] / tagged[("Cohort", "Unfavourable")] / (
        good_flagged_bad["Old"] / tagged[("Old", "Unfavourable")])
    if r_missed < 0.9:
        band_runs += [(", with fewer good clients wrongly tagged Unfavourable (", {}),
                      (f"{pct(good_flagged_bad['Cohort'] / tagged[('Cohort', 'Unfavourable')])} vs "
                       f"{pct(good_flagged_bad['Old'] / tagged[('Old', 'Unfavourable')])}", {"bold": True, "color": GOLD}),
                      (")", {})]
    band_runs.append((".", {}))
    impact_band(s, "WHAT IT MEANS", band_runs, size=11.5)
    footer(s, page0 + 2)
    return 3 + n_dd


# ---------- volume view (FundedLoanAmt) and win-rate implication ----------
FUNDED_VIEWS = RESULTS.parent / "b2b_funded_amount_views.xlsx"
fv = read_xl(FUNDED_VIEWS, sheet_name=None)
fmix = fv["Classification Mix"].set_index(["Population", "Approach", "Classification"])
ftab = fv["Funded Cross Tabs"]

vol_tot = scores["funded"].sum()
vol = {a: scores.groupby(f"{k}_band")["funded"].sum().reindex(BANDS, fill_value=0.0)
       for a, k in [("Old", "old"), ("Cohort", "new")]}
cnt = {a: scores[f"{k}_band"].value_counts().reindex(BANDS, fill_value=0) for a, k in [("Old", "old"), ("Cohort", "new")]}
vshare = {a: vol[a] / vol_tot for a in APPROACHES}
vxt = (scores.pivot_table(index="old_band", columns="new_band", values="funded", aggfunc="sum")
       .reindex(index=BANDS, columns=BANDS).fillna(0.0))                       # rows Old, cols Cohort
# tie-out: volume mix and volume cross-tab reproduce the notebook's funded views (Section 11C)
for a in APPROACHES:
    for b in BANDS:
        assert abs(vshare[a][b] - fmix.loc[("Original Base", a, b), "Funded %"]) < 1e-4, (a, b)
_nb_x = ftab[ftab["View"].eq("Cohort vs Old: Original Base") & ftab["Weight basis"].eq("Cohort funding")]
_nb_x = _nb_x.pivot(index="Column classification", columns="Row classification", values="FundedLoanAmt")
assert np.allclose(_nb_x.reindex(index=BANDS, columns=BANDS).values, vxt.values, rtol=1e-6)
vsame = sum(vxt.loc[b, b] for b in BANDS) / vol_tot
vopp = (vxt.loc["Favourable", "Unfavourable"] + vxt.loc["Unfavourable", "Favourable"]) / vol_tot
vone = 1 - vsame - vopp

# forward test weighted by base-period funded volume (notebook Section 6 funded transitions)
fwd_v = {}
for a in APPROACHES:
    sub = ftab[ftab["View"].str.startswith(f"{a} Forward")]
    fwd_v[a] = sub.pivot(index="Row classification", columns="Column classification",
                         values="FundedLoanAmt").reindex(index=BANDS, columns=BANDS).fillna(0.0)
v_tagged = {(a, b): fwd_v[a].loc[b].sum() for a in APPROACHES for b in BANDS}
v_hit = {(a, b): fwd_v[a].loc[b, b] for a in APPROACHES for b in BANDS}
v_hit_rate = {k: v_hit[k] / v_tagged[k] for k in v_hit}
v_tag_rate = {a: (v_hit[(a, "Favourable")] + v_hit[(a, "Unfavourable")])
              / (v_tagged[(a, "Favourable")] + v_tagged[(a, "Unfavourable")]) for a in APPROACHES}
v_missed = {a: fwd_v[a].loc["Unfavourable", "Favourable"] / v_tagged[(a, "Unfavourable")] for a in APPROACHES}

# win-rate implication: total purchases held at today's win rate, bought in the target tag mix
WIN_TODAY = 0.03
TARGET_MIX = {"Favourable": 0.85, "Neutral": 0.10, "Unfavourable": 0.05}
need = {(a, b): TARGET_MIX[b] * WIN_TODAY / vshare[a][b] for a in APPROACHES for b in BANDS}
buy_tot = WIN_TODAY * vol_tot
win_windows = [("Main comparison", {a: vshare[a]["Favourable"] for a in APPROACHES})]
for lab, pop in [(F_BASE, NB["base_period"]), (F_COMP, NB["comp_period"])]:
    win_windows.append((lab, {a: fmix.loc[(pop, a, "Favourable"), "Funded %"] for a in APPROACHES}))


def bn(v, d=1):
    """$ in billions, or millions below $1B."""
    return f"${v / 1e9:,.{d}f}B" if abs(v) >= 1e9 else f"${v / 1e6:,.0f}M"


def volume_slides(prs, page0):
    """Slides: volume view of tags, head-to-head and stability; win-rate implication. Returns slide count."""

    # =====================================================================
    # VOLUME VIEW — tag mix, head-to-head and stability weighted by FundedLoanAmt
    # =====================================================================
    s = new_slide(prs)
    tag(s, "VOLUME VIEW", INDIGO)
    slide_title(
        s,
        f"Weighted by funded volume, the approaches agree on {pct(vsame)} of dollars; Cohort tags "
        f"{pct(vshare['Cohort']['Unfavourable'])} of volume Unfavourable against {pct(vshare['Old']['Unfavourable'])}",
        f"Funding {FUND_WINDOW} · {bn(vol_tot, 2)} across {n_clients} clients · "
        "Same TotalScore tags, weighted by client funded volume",
    )

    # ----- Panel 1: volume by tag (strategy-deck summary format) -----
    P1X, P1W = 0.5, 4.3
    box(s, P1X, PY, P1W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P1X + PAD, P1W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 1, "VOLUME BY TAG",
                 f"Favourable: {pct(vshare['Cohort']['Favourable'])} of volume under Cohort, "
                 f"{pct(vshare['Old']['Favourable'])} under Old")
    by = PY + 1.02
    text(s, ix, by, iw, 0.2, [[("Share of funded volume   ", {"bold": True, "color": INK, "size": 9.5}),
                               ("■ ", {"color": GREEN}), ("Fav.  ", {}), ("■ ", {"color": REFGRAY}), ("Neutral  ", {}),
                               ("■ ", {"color": RED}), ("Unfav.", {})]], size=8.5, color=MUTE)
    lw2 = 0.62
    bx0, bw_ = ix + lw2, iw - lw2
    for k, (a, col) in enumerate([("Old", OLD_C), ("Cohort", INDIGO)]):
        ry = by + 0.26 + k * 0.38
        text(s, ix, ry, lw2 - 0.06, 0.3, a, size=9.5, bold=True, color=col, valign="middle")
        sx = bx0
        for b, fill, fc in [("Favourable", GREEN, WHITE), ("Neutral", REFGRAY, INK), ("Unfavourable", RED, WHITE)]:
            sw = bw_ * vshare[a][b]
            box(s, sx, ry, sw, 0.3, fill=fill)
            if sw > 0.32:
                text(s, sx, ry, sw, 0.3, pct(vshare[a][b]), size=9.5, bold=True, color=fc, align="center",
                     valign="middle")
            sx += sw

    ty = by + 1.12
    tcols = [("", 0.95, "left"), ("Volume", 0.56, "right"), ("Share", 0.44, "right"), ("Clients", 0.45, "right"),
             ("Volume", 0.56, "right"), ("Share", 0.44, "right"), ("Clients", 0.45, "right")]
    xs, cx_ = [], ix
    for k, (_, w_, _) in enumerate(tcols):
        xs.append(cx_)
        cx_ += w_ + (0.0 if k != 3 else 0.0)
    for (lab, col, c0) in [("OLD", OLD_C, 1), ("COHORT", INDIGO, 4)]:
        gx0, gx1 = xs[c0], xs[c0 + 2] + tcols[c0 + 2][1]
        text(s, gx0, ty, gx1 - gx0, 0.17, lab, size=8, bold=True, color=col, align="center", char_spacing=1.5)
        line(s, gx0 + 0.05, ty + 0.2, gx1, ty + 0.2, col, 1)
    for (lab, w_, al), x_ in zip(tcols, xs):
        text(s, x_, ty + 0.22, w_, 0.17, lab, size=7.5, bold=True, color=MUTE, align=al)
    trh = 0.25
    rows_ = [(b, vol["Old"][b], vshare["Old"][b], cnt["Old"][b], vol["Cohort"][b], vshare["Cohort"][b], cnt["Cohort"][b])
             for b in BANDS] + [("Total", vol_tot, 1.0, n_clients, vol_tot, 1.0, n_clients)]
    for k, (b, ov, osh, oc, cv, csh, cc) in enumerate(rows_):
        yy = ty + 0.42 + k * trh
        tot = b == "Total"
        if tot:
            line(s, ix, yy, ix + iw, yy, MUTE, 0.75)
        elif k % 2 == 0:
            box(s, ix - 0.05, yy, iw + 0.1, trh, fill=WHITE)
        bcol = TAG_COL[b][0] if b in TAG_COL else INK
        vals = [(b, {"bold": True, "color": bcol}), (bn(ov), {}), (pct(osh, 1), {"bold": True, "color": OLD_C}),
                (f"{oc:,}", {"color": MUTE}), (bn(cv), {}), (pct(csh, 1), {"bold": True, "color": INDIGO}),
                (f"{cc:,}", {"color": MUTE})]
        for ci, (txt, o) in enumerate(vals):
            text(s, xs[ci], yy, tcols[ci][1], trh, txt, size=8.5, color=o.get("color", INK),
                 bold=o.get("bold", tot), align=tcols[ci][2], valign="middle")
    d_unf = vol["Cohort"]["Unfavourable"] - vol["Old"]["Unfavourable"]
    ny = ty + 0.42 + 4 * trh + 0.08
    text(s, ix, ny, iw, PY + PH - 0.08 - ny, [[
        (f"Cohort tags {bn(abs(d_unf))} {'more' if d_unf > 0 else 'less'} volume Unfavourable", {"bold": True}),
        (f": {pct(cnt['Cohort']['Unfavourable'] / n_clients)} of clients but "
         f"{pct(vshare['Cohort']['Unfavourable'])} of volume, so the Unfavourable group is "
         f"{'larger' if vshare['Cohort']['Unfavourable'] > cnt['Cohort']['Unfavourable'] / n_clients else 'smaller'}"
         " clients on average.", {}),
    ]], size=8.5, color=INK, spacing=0.95)

    # ----- Panel 2: head-to-head by volume (emphasis) -----
    P2X, P2W = P1X + P1W + 0.2, 4.35
    box(s, P2X, PY, P2W, PH, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = P2X + PAD, P2W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 2, "HEAD-TO-HEAD BY VOLUME",
                 f"{pct(vsame)} of volume keeps the same tag")
    text(s, ix, PY + 0.8, iw, 0.2, [[("Rows: ", {}), ("Old", {"bold": True, "color": OLD_C}),
         (" tag   ·   Columns: ", {}), ("Cohort", {"bold": True, "color": INDIGO}),
         (" tag   ·   % of total volume", {})]], size=9, color=MUTE)
    gy = PY + 1.04
    lab_w, rh, hh = 0.7, 0.56, 0.24
    cw = (iw - lab_w) / 3
    short = {"Favourable": "Fav.", "Neutral": "Neutral", "Unfavourable": "Unfav."}
    for j, b in enumerate(BANDS):
        text(s, ix + lab_w + j * cw, gy, cw, hh, short[b], size=9.5, bold=True, color=INDIGO,
             align="center", valign="middle")
    for i, rb in enumerate(BANDS):
        ry = gy + hh + i * rh
        text(s, ix, ry, lab_w, rh, short[rb], size=9.5, bold=True, color=OLD_C, valign="middle")
        for j, cb in enumerate(BANDS):
            v = vxt.loc[rb, cb]
            sh = v / vol_tot
            if i == j:
                fill, fc, sc = INDIGO, WHITE, INDIGO_TEXT_LT
            elif abs(i - j) == 1:
                fill, fc, sc = (AMBER_TINT, AMBER, AMBER) if sh >= 0.05 else (WHITE, MUTE, MUTE)
            else:
                fill, fc, sc = (RED_TINT, RED, RED) if v else (WHITE, MUTE, MUTE)
            box(s, ix + lab_w + j * cw + 0.03, ry + 0.03, cw - 0.06, rh - 0.06, fill=fill,
                line=LINEGRAY if fill == WHITE else None, line_w=0.75, radius=0.05)
            text(s, ix + lab_w + j * cw, ry + 0.02, cw, rh - 0.04, [
                [(pct(sh, 1), {"size": 13, "bold": True, "color": fc})],
                [(bn(v), {"size": 8, "color": sc})],
            ], align="center", valign="middle", spacing=0.85)
    ly = gy + hh + 3 * rh + 0.08
    text(s, ix, ly, iw, 0.22, [[
        ("■ ", {"color": INDIGO}), (f"Same {pct(vsame)}", {"bold": True}),
        ("   ■ ", {"color": AMBER}), (f"One step {pct(vone)}", {}),
        ("   ■ ", {"color": RED}), (f"Opposite {pct(vopp, 1)}", {}),
    ]], size=9.5, color=INK)
    off = sorted(((vxt.loc[r, c], r, c) for r in BANDS for c in BANDS if r != c), reverse=True)
    v1, r1, c1 = off[0]
    v2, r2, c2 = off[1]
    text(s, ix, ly + 0.3, iw, PY + PH - 0.08 - ly - 0.3, [
        [(f"Same tag on {pct(vsame)} of volume but {pct(same / n_clients)} of clients: ", {"bold": True}),
         ("the larger clients agree more often.", {})],
        [(f"Largest shift: Old {short[r1]} → Cohort {short[c1]}, {bn(v1)} ({pct(v1 / vol_tot)} of volume).",
          {"bold": True, "space_before": 4})],
        [(f"Next: Old {short[r2]} → Cohort {short[c2]}, {bn(v2)} ({pct(v2 / vol_tot)}). "
          f"Opposite reversals are {bn(vopp * vol_tot)}, led by {short_name(flip_leader)}.",
          {"space_before": 4})],
    ], size=9, color=INK, spacing=0.96)

    # ----- Panel 3: stability by volume -----
    P3X = P2X + P2W + 0.2
    P3W = W - 0.5 - P3X
    box(s, P3X, PY, P3W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P3X + PAD, P3W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 3, "STABILITY BY VOLUME",
                 f"Share of {F_BASE} Fav. / Unfav. volume confirmed in {F_COMP}")
    text(s, ix, PY + 0.86, iw, 0.5, [[
        (pct(v_tag_rate["Cohort"], 1), {"size": 26, "bold": True, "color": INDIGO}),
        ("  vs  ", {"size": 11, "color": MUTE}),
        (pct(v_tag_rate["Old"], 1), {"size": 26, "bold": True, "color": OLD_C}),
    ]], valign="middle")
    text(s, ix, PY + 1.36, iw, 0.2, [[
        (f"{bn(v_hit[('Cohort', 'Favourable')] + v_hit[('Cohort', 'Unfavourable')])} of "
         f"{bn(v_tagged[('Cohort', 'Favourable')] + v_tagged[('Cohort', 'Unfavourable')])}", {"bold": True, "color": INDIGO}),
        ("  vs  ", {}),
        (f"{bn(v_hit[('Old', 'Favourable')] + v_hit[('Old', 'Unfavourable')])} of "
         f"{bn(v_tagged[('Old', 'Favourable')] + v_tagged[('Old', 'Unfavourable')])}", {"bold": True, "color": OLD_C}),
    ]], size=8.5, color=MUTE)
    text(s, ix, PY + 1.64, iw, 0.2, [[
        ("■ ", {"color": GREEN}), ("Confirmed  ", {}), ("■ ", {"color": REFGRAY}), ("Neutral  ", {}),
        ("■ ", {"color": AMBER}), ("Reversed", {}),
    ]], size=8.5, color=MUTE)
    rows = [("Favourable", "Cohort"), ("Favourable", "Old"), ("Unfavourable", "Cohort"), ("Unfavourable", "Old")]
    bar_w_max = iw                                             # every bar = 100% of that tag's volume
    for r, (band, a) in enumerate(rows):
        yy = PY + 1.9 + r * 0.44 + (0.08 if r >= 2 else 0)
        col = INDIGO if a == "Cohort" else OLD_C
        opp = "Unfavourable" if band == "Favourable" else "Favourable"
        text(s, ix, yy, iw, 0.17, [[
            (f"{short[band]} · ", {"color": INK}), (a, {"bold": True, "color": col}),
            (f"  {bn(v_tagged[(a, band)])} tagged", {"color": MUTE}),
        ]], size=8.5, valign="middle")
        segs = [(v_hit[(a, band)], GREEN, WHITE), (fwd_v[a].loc[band, "Neutral"], REFGRAY, INK),
                (fwd_v[a].loc[band, opp], AMBER, WHITE)]
        # tie-out: segment shares equal the notebook's Row % for this base tag
        nb_row = ftab[ftab["View"].str.startswith(f"{a} Forward") & ftab["Row classification"].eq(band)]
        nb_row = nb_row.set_index("Column classification")["Row %"]
        assert all(abs(v / v_tagged[(a, band)] - nb_row[c]) < 1e-6
                   for (v, _, _), c in zip(segs, [band, "Neutral", opp])), (a, band)
        sx = ix
        for v, fill, fc in segs:
            sw = bar_w_max * v / v_tagged[(a, band)]
            if sw <= 0:
                continue
            box(s, sx, yy + 0.2, sw, 0.2, fill=fill)
            if sw > 0.17:
                text(s, sx, yy + 0.2, sw, 0.2, pct(v / v_tagged[(a, band)]), size=8, bold=True, color=fc,
                     align="center", valign="middle")
            sx += sw
    cy_ = PY + 1.9 + 4 * 0.44 + 0.14
    text(s, ix, cy_, iw, PY + PH - 0.06 - cy_, [[
        ("Good clients missed: ", {"bold": True}),
        (f"{pct(v_missed['Cohort'])} vs {pct(v_missed['Old'])}", {"bold": True, "color": INDIGO}),
        (" of Unfavourable volume turned Favourable.", {}),
    ]], size=8.5, color=INK, spacing=0.95)
    impact_band(s, "WHAT IT MEANS", [
        (f"{pct(vsame)} of funded volume keeps the same tag. ", {"bold": True, "color": WHITE}),
        (f"Forward volume retention: Cohort {pct(v_tag_rate['Cohort'], 1)} vs Old {pct(v_tag_rate['Old'], 1)}.",
         {"color": GOLD}),
    ])
    footer(s, page0)

    # =====================================================================
    # WIN RATE — what the target buy mix implies for Favourable win rate
    # =====================================================================
    s = new_slide(prs)
    tag(s, "WIN RATE", INDIGO)
    fav_c, fav_o = need[("Cohort", "Favourable")], need[("Old", "Favourable")]
    mix_txt = " / ".join(f"{TARGET_MIX[b] * 100:.0f}" for b in BANDS)
    slide_title(
        s,
        f"To buy {pct(TARGET_MIX['Favourable'])} Favourable at today’s ~{pct(WIN_TODAY)} win rate, we need to win "
        f"{pct(fav_c, 1)} of Favourable volume on Cohort tags, against {pct(fav_o, 1)} on today’s Old tags",
        f"Illustrative. Last-12-month FundedLoanAmt ({bn(vol_tot)}) stands in for the market; purchases held at "
        f"{pct(WIN_TODAY)} ({bn(buy_tot, 2)}). Win rate = $ bought ÷ $ available in the tag.",
    )

    # ----- Panel 1: the arithmetic, step by step (Cohort tags) -----
    P1X, P1W = 0.5, 3.95
    box(s, P1X, PY, P1W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P1X + PAD, P1W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 1, "THE ARITHMETIC",
                 f"Win rate needed = target share × {pct(WIN_TODAY)} ÷ market share")
    sv = vshare["Cohort"]
    steps = [
        ("Market volume by Cohort tag",
         [[[(pct(sv[b]), {"bold": True, "color": TAG_COL[b][0], "size": 11})], [(f"{short[b]} {bn(vol['Cohort'][b])}", {})]]
          for b in BANDS]),
        (f"Purchases at today’s {pct(WIN_TODAY)} win rate",
         [[(f"{bn(vol_tot)} × {pct(WIN_TODAY)} = ", {}), (bn(buy_tot, 2), {"bold": True, "color": INDIGO}),
           ("   12-month funded volume", {"size": 8, "color": MUTE})]]),
        (f"Target buy mix {mix_txt}",
         [[[(bn(TARGET_MIX[b] * buy_tot, 2), {"bold": True, "color": TAG_COL[b][0], "size": 11})],
           [(f"{short[b]} {pct(TARGET_MIX[b])}", {})]] for b in BANDS]),
        ("Win rate needed on Favourable",
         [[(f"{bn(TARGET_MIX['Favourable'] * buy_tot, 2)} ÷ {bn(vol['Cohort']['Favourable'])} = ", {}),
           (pct(fav_c, 1), {"bold": True, "color": INDIGO, "size": 12})]]),
    ]
    sy0, sgap = PY + 1.02, 0.08
    sh_ = (PY + PH - 0.14 - sy0 - 3 * sgap) / 4
    for k, (ttl, body) in enumerate(steps):
        yy = sy0 + k * (sh_ + sgap)
        box(s, ix, yy, iw, sh_, fill=WHITE if k < 3 else INDIGO_TINT, radius=0.06,
            line=INDIGO if k == 3 else None, line_w=1.0)
        box(s, ix + 0.1, yy + 0.1, 0.26, 0.26, fill=INDIGO, radius=0.13)
        text(s, ix + 0.1, yy + 0.1, 0.26, 0.26, str(k + 1), size=9.5, bold=True, color=WHITE, align="center",
             valign="middle")
        text(s, ix + 0.46, yy + 0.06, iw - 0.56, 0.34, ttl, size=9.5, bold=True, color=INK, valign="middle")
        if len(body) == 3:                                         # three tags side by side
            bw3 = (iw - 0.56) / 3
            for j, paras in enumerate(body):
                text(s, ix + 0.46 + j * bw3, yy + 0.36, bw3, sh_ - 0.4, paras, size=8.5, color=INK,
                     valign="middle", spacing=0.9)
        else:
            text(s, ix + 0.46, yy + 0.4, iw - 0.56, sh_ - 0.46, body, size=9.5, color=INK, valign="middle")

    # ----- Panel 2: win rate needed, by tag, Old vs Cohort (emphasis) -----
    P2X, P2W = P1X + P1W + 0.2, 4.95
    box(s, P2X, PY, P2W, PH, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = P2X + PAD, P2W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 2, "WIN RATE NEEDED BY TAG",
                 f"Favourable roughly {'doubles' if 1.8 <= fav_c / WIN_TODAY < 2.3 else 'rises'}; "
                 f"Neutral and Unfavourable fall below {pct(WIN_TODAY)}")
    text(s, ix, PY + 0.98, iw, 0.2, [[
        ("■ ", {"color": OLD_C}), ("Old tags (today)   ", {}), ("■ ", {"color": INDIGO}), ("Cohort tags   ", {}),
        ("┆ ", {"color": AMBER, "bold": True}), (f"today’s ~{pct(WIN_TODAY)} on all volume", {}),
    ]], size=8.5, color=MUTE)
    lab_w = 1.55
    ax0, ax1 = ix + lab_w, ix + iw - 0.55
    vmax = 0.07 if max(need.values()) < 0.07 else float(np.ceil(max(need.values()) * 100 + 1) / 100)
    xpos = lambda v: ax0 + (ax1 - ax0) * v / vmax
    gy0, gh = PY + 1.3, 0.78
    bar_h = 0.24
    for k, b in enumerate(BANDS):
        yy = gy0 + k * gh
        text(s, ix, yy, lab_w - 0.1, 0.3, [[(b, {"bold": True, "color": TAG_COL[b][0]})]], size=10, valign="middle")
        text(s, ix, yy + 0.27, lab_w - 0.1, 0.4, [
            [(f"target {pct(TARGET_MIX[b])} of buys", {})],
            [("pool ", {}), (bn(vol["Old"][b]), {"color": OLD_C, "bold": True}), (" / ", {}),
             (bn(vol["Cohort"][b]), {"color": INDIGO, "bold": True})],
        ], size=8, color=MUTE, spacing=0.9)
        for j, (a, col) in enumerate([("Old", OLD_C), ("Cohort", INDIGO)]):
            v = need[(a, b)]
            by_ = yy + 0.04 + j * (bar_h + 0.06)
            box(s, ax0, by_, max(xpos(v) - ax0, 0.02), bar_h, fill=col)
            text(s, xpos(v) + 0.05, by_ - 0.02, 0.55, bar_h + 0.04, pct(v, 1), size=10.5 if b == "Favourable" else 9.5,
                 bold=True, color=col, valign="middle")
    axy = gy0 + 3 * gh - 0.16
    line(s, ax0, axy, ax1, axy, REFGRAY, 1)
    for v in np.arange(0, vmax + 1e-9, 0.01):
        text(s, xpos(v) - 0.25, axy + 0.03, 0.5, 0.15, f"{v * 100:.0f}%", size=7.5, color=MUTE, align="center")
    dashed(s, xpos(WIN_TODAY), gy0 - 0.06, xpos(WIN_TODAY), axy, AMBER, 1.5)
    nt = {a: vshare[a] for a in APPROACHES}
    ny = axy + 0.26
    text(s, ix, ny, iw, PY + PH - 0.08 - ny, [[
        (f"Winning {pct(WIN_TODAY)} everywhere buys the market mix: ", {"bold": True}),
        (" / ".join(f"{nt['Cohort'][b] * 100:.0f}" for b in BANDS), {"bold": True, "color": INDIGO}),
        (" on Cohort tags, ", {}),
        (" / ".join(f"{nt['Old'][b] * 100:.0f}" for b in BANDS), {"bold": True, "color": OLD_C}),
        (f" on Old, against the {mix_txt} target.", {}),
    ]], size=9, color=INK, spacing=0.95)

    # ----- Panel 3: sensitivity -----
    P3X = P2X + P2W + 0.2
    P3W = W - 0.5 - P3X
    box(s, P3X, PY, P3W, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = P3X + PAD, P3W - 2 * PAD
    panel_header(s, ix, PY + 0.16, iw, 3, "SENSITIVITY", "Favourable win rate needed, other assumptions")
    c0w = iw - 2 * 0.78
    tcols = [(c0w, "left"), (0.78, "center"), (0.78, "center")]
    txs = [ix, ix + c0w, ix + c0w + 0.78]

    def table(y0, head, rows_, hl=None):
        text(s, ix, y0, iw, 0.2, head, size=8.5, bold=True, color=INK, valign="middle")
        for (lab, col), x_, (w_, al) in zip([("", MUTE), ("Old", OLD_C), ("Cohort", INDIGO)], txs, tcols):
            text(s, x_, y0 + 0.22, w_, 0.17, lab, size=7.5, bold=True, color=col, align=al)
        for k, (lab, o_, c_) in enumerate(rows_):
            yy = y0 + 0.4 + k * 0.23
            hi = k == hl
            if hi:
                box(s, ix - 0.05, yy, iw + 0.1, 0.23, fill=INDIGO_TINT, radius=0.04)
            elif k % 2 == 0:
                box(s, ix - 0.05, yy, iw + 0.1, 0.23, fill=WHITE)
            for (txt, col, bold_), x_, (w_, al) in zip([(lab, INK, hi), (o_, OLD_C, True), (c_, INDIGO, True)], txs, tcols):
                text(s, x_ + (0.04 if al == "left" else 0), yy, w_, 0.23, txt, size=8.5, color=col, bold=bold_,
                     align=al, valign="middle")
        return y0 + 0.4 + len(rows_) * 0.23

    wr_rows = [(f"{pct(w_)} overall" + ("  (today)" if w_ == WIN_TODAY else ""),
                pct(TARGET_MIX["Favourable"] * w_ / vshare["Old"]["Favourable"], 1),
                pct(TARGET_MIX["Favourable"] * w_ / vshare["Cohort"]["Favourable"], 1))
               for w_ in (0.02, 0.03, 0.04, 0.05)]
    y_ = table(PY + 1.02, "If the overall win rate changes", wr_rows, hl=1)
    win_rows = [(lab, pct(TARGET_MIX["Favourable"] * WIN_TODAY / sh["Old"], 1),
                 pct(TARGET_MIX["Favourable"] * WIN_TODAY / sh["Cohort"], 1)) for lab, sh in win_windows]
    y_ = table(y_ + 0.14, "If the market’s tag mix looks like another window", win_rows, hl=0)
    text(s, ix, y_ + 0.08, iw, PY + PH - 0.06 - y_ - 0.08, [[
        ("Cohort needs a higher Favourable win rate ", {"bold": True}),
        ("because it tags less volume Favourable; the gap holds in every window.", {}),
    ]] if all(sh["Cohort"] < sh["Old"] for _, sh in win_windows) else [[
        ("The gap between approaches depends on the window.", {"bold": True})]], size=8.5, color=INK, spacing=0.95)
    impact_band(s, "IMPLICATION", [
        (f"Illustrative {mix_txt} buy mix at {pct(WIN_TODAY)} overall win rate: ",
         {"bold": True, "color": WHITE}),
        (f"win {pct(fav_c, 1)} of Cohort Favourable volume vs {pct(fav_o, 1)} under Old.", {"color": GOLD}),
    ])
    footer(s, page0 + 1)
    return 2


def appendix_slides(prs, page0):
    """Appendix: reversals and top 25 lenders, full score detail. Returns slide count."""
    # =====================================================================
    # APPENDIX — reversals and top 25 lenders, full score detail
    # =====================================================================
    s = new_slide(prs)
    tag(s, "APPENDIX", OLD_C)
    if DECK_MODE["compact"]:
        STRETCH[s.slide_id] = (PY, 5.95, CONTENT_BOTTOM)
    slide_title(
        s,
        f"Only {opposite} clients reverse between Favourable and Unfavourable",
        f"Funding {FUND_WINDOW} · Old portfolio score {Z['Old']:.2f}; Cohort {Z['Cohort']:.2f} · "
        "Reversals and top lenders by volume",
    )
    tx0, tw0 = 0.5, W - 1.0
    box(s, tx0, PY, tw0, PH, fill=GRAYFILL, radius=0.08)
    ix, iw = tx0 + PAD, tw0 - 2 * PAD
    tcols = [("#", 0.4, "left"), ("Client", 2.9, "left"), ("Funded", 0.85, "right"),
             ("CPR score", 0.8, "center"), ("Recapture", 0.8, "center"), ("TotalScore", 0.85, "center"),
             ("Tag", 0.45, "center"),
             ("CPR score", 0.8, "center"), ("Recapture", 0.8, "center"), ("TotalScore", 0.85, "center"),
             ("Tag", 0.45, "center"), ("Note", 0.0, "left")]
    widths = [w_ for _, w_, _ in tcols[:-1]]
    widths.append(iw - sum(widths) - 0.2 * 2)
    xs, cx_ = [], ix
    for k, w_ in enumerate(widths):
        xs.append(cx_)
        cx_ += w_ + (0.2 if k in (2, 6) else 0)
    hy = PY + 0.1
    for (span0, span1, lab, col) in [(3, 6, "OLD", OLD_C), (7, 10, "COHORT", INDIGO)]:
        gx0, gx1 = xs[span0], xs[span1] + widths[span1]
        text(s, gx0, hy, gx1 - gx0, 0.18, lab, size=8, bold=True, color=col, align="center", char_spacing=1.5)
        line(s, gx0, hy + 0.21, gx1, hy + 0.21, col, 1)
    for (lab, _, al), x_, w_ in zip(tcols, xs, widths):
        text(s, x_, hy + 0.24, w_, 0.17, lab, size=7.5, bold=True, color=MUTE, align=al)

    def detail(names):
        """One row per client: rank, funded, and both approaches' CPR / recapture / TotalScore / tag."""
        rows = []
        for n in names:
            sc = scores.loc[n]
            note = ("No prepayments yet: default score" if sc["zero_prepay"] else
                    {1: "Better tag under Cohort", -1: "Lower tag under Cohort", 0: ""}[int(sc["move"])])
            rows.append((int(sc["rank"]), n, sc["funded"],
                         old_clients.loc[n, "CPRScore"], old_clients.loc[n, "RecapScore"], sc["old"], sc["old_band"],
                         cohort_clients.loc[n, "CPRScore"], cohort_clients.loc[n, "RecapScore"], sc["new"],
                         sc["new_band"], note))
        return rows

    rev_names = {k: v.sort_values(["ComparisonVolumeRank", "MetricValue"])["MetricValue"].tolist() for k, v in rev.items()}
    sections = [("Reversals · Old Unfavourable → Cohort Favourable", detail(rev_names["UF"])),
                ("Reversals · Old Favourable → Cohort Unfavourable", detail(rev_names["FU"])),
                (f"Top 25 lenders by volume · {pct(top25_share)} of funded volume", detail(top25.index.tolist()))]
    n_lines = sum(len(r) + 1 for _, r in sections)
    yy = hy + 0.45
    rh = min(0.13, (PY + PH - 0.08 - yy - 0.05 * len(sections)) / n_lines)
    fs = 7.5 if rh >= 0.12 else 7
    for title, rows in sections:
        text(s, ix, yy, iw, rh, [[(title, {"bold": True, "color": INK}), (f"   {len(rows)} clients", {"color": MUTE})]],
             size=fs + 0.5, valign="middle")
        yy += rh
        for k, r in enumerate(rows):
            if k % 2 == 0:
                box(s, ix - 0.05, yy, iw + 0.1, rh, fill=WHITE)
            rank, name, funded, ocpr, orec, ots, oband, ccpr, crec, cts, cband, note = r
            vals = [(str(rank), {"color": MUTE}), (short_name(name), {"bold": name in DEEP_DIVE_CLIENTS}),
                    (f"${funded / 1e6:,.1f}M", {}),
                    (f"{ocpr:.2f}", {}), (f"{orec:.2f}", {}), (f"{ots:.2f}", {"bold": True, "color": OLD_C}), oband,
                    (f"{ccpr:.2f}", {}), (f"{crec:.2f}", {}), (f"{cts:.2f}", {"bold": True, "color": INDIGO}), cband,
                    (note, {"color": MUTE})]
            for ci, v in enumerate(vals):
                if isinstance(v, str):
                    tag_chip(s, xs[ci] + (widths[ci] - 0.26) / 2, yy + 0.008, 0.26, rh - 0.016, v, size=6.5)
                    continue
                txt, o = v
                text(s, xs[ci], yy, widths[ci], rh, txt, size=fs, color=o.get("color", INK),
                     bold=o.get("bold", False), align=tcols[ci][2], valign="middle")
            yy += rh
        yy += 0.05
    assert yy < PY + PH, "appendix table overflows its panel"
    fwd = " · ".join(f"{w.replace('-', '–')} {a} U→F, {b} F→U" for w, (a, b) in fwd_rev.items())
    note_runs = []
    if rev_zero["FU"]:
        note_runs.append((f"{rev_zero['FU']} of the {fav_to_unf} Old F → Cohort U clients "
                          + ("has" if rev_zero["FU"] == 1 else "have") + " no prepayments yet (default score). ",
                          {"bold": True, "color": INK if DECK_MODE["compact"] else WHITE}))
    note_runs.append(("F / N / U = Favourable / Neutral / Unfavourable."
                      + (f"  Forward-test reversals: {fwd}." if fwd else ""), {}))
    impact_band(s, "READING NOTE", note_runs, size=10.5)
    footer(s, page0)
    return 1


# ---------- cohort facts (from the May 2026 cohort scoring run) ----------
cohort_ids = inc["Cohort_ID"].drop_duplicates()
BUILD = {"cohorts": int(cohort_ids.size)}
mapped_share = inc["Client_LoanCt"].sum() / audit["Client_LoanCt"].sum()
loan_months = int(audit["Client_LoanCt"].sum())
scored = inc.drop_duplicates("Cohort_ID")
cohort_cpr_q = scored["Cohort_x_Cohort_ActualCPR"].quantile([0.25, 0.75, 1.0]).tolist()
lt_all = inc["Rule_bucket"].str.extract(r"LoanType=(\w+)")[0]
type_share = (inc.groupby(lt_all)["Client_BeginPrinBal"].sum() / inc["Client_BeginPrinBal"].sum()).to_dict()
by_cohort = inc.groupby(["Cohort_ID", "Rule_bucket"])["Client_BeginPrinBal"].sum().sort_values(ascending=False)
(EX_COHORT, EX_RULE), ex_bal = by_cohort.index[0], by_cohort.iloc[0]
ex_cohort_share = ex_bal / inc["Client_BeginPrinBal"].sum()
edges = sorted({float(v) for v in inc["Rule_bucket"].str.findall(r"([+-]?\d+(?:\.\d+)?)%").sum()})


def describe_rule(rule):
    """'LoanType=CONV | FICOOrig >750 | ...' -> exec-readable cohort description."""
    out = []
    for part in [p.strip() for p in rule.split("|")]:
        if part.startswith("LoanType="):
            out.append({"CONV": "Conventional"}.get(part.split("=")[1], part.split("=")[1]))
        elif part.startswith("FICOOrig"):
            out.append("FICO " + part.split()[1].replace("<=", "≤ ").replace(">", "> "))
        elif part.startswith("RemainingTerm"):
            out.append("> 20 yr term" if ">" in part else "≤ 20 yr term")
        elif part.startswith("PrinBal"):
            out.append(part.replace("PrinBal ", "").replace("(-Inf, ", "≤ $").replace("]", "")
                       .replace(">", "> $").replace("k", "K"))
        elif part.startswith("RateSpread"):
            lo, hi = [float(v) for v in re.findall(r"([+-]?(?:[\d.]+|Inf))", part.split(" ", 1)[1])[:2]]
            fmt = lambda v: f"{abs(v):g}"
            out.append(f"rate {fmt(hi)}–{fmt(lo)}% below market" if hi <= 0 else
                       f"rate {fmt(lo)}–{fmt(hi)}% above market" if lo >= 0 else "rate near market")
    return " · ".join(out)


EX = "LOWER LLC"
ex = deep[EX]


def k_fmt(n):
    return f"{n / 1e6:.1f}M" if n >= 1e6 else f"{n / 1e3:,.1f}K"


def methodology_slides(prs, page0):
    """Slides: overview (what / why / how) and cohort membership & definitions. Returns slide count."""
    # =====================================================================
    # OVERVIEW — the proposal, and how the score is produced
    # =====================================================================
    s = new_slide(prs)
    tag(s, "OVERVIEW", INDIGO)
    STRETCH.pop(s.slide_id, None)          # laid out directly to CONTENT_BOTTOM
    slide_title(
        s,
        "We propose scoring both prepayment and recapture against actual performance of look-alike loans "
        "rather than against model predictions",
        "Proposed benchmarks for each part of the client score; below, how the prepayment score is produced "
        "for one client.",
    )
    proposal = [
        ("PREPAYMENT", [
            ("V1", "Benchmark actual CPR against monthly modeled lifetime CPR"),
            ("V2", "Benchmark actual CPR against observed performance across 100+ valuation model cohorts"),
        ]),
        ("RECAPTURE", [
            ("V1", "Benchmark actual recapture against observed performance by product, payoff reason and "
                   "payoff month"),
            ("V2", "Benchmark actual recapture against observed performance at the overall portfolio level, "
                   "payoff reason and payoff month"),
        ]),
    ]
    cy, chh, cg = PY, 1.72, 0.2
    cwid = (W - 1.0 - cg) / 2
    for k, (head, items) in enumerate(proposal):
        cx = 0.5 + k * (cwid + cg)
        box(s, cx, cy, cwid, chh, fill=GRAYFILL, radius=0.08)
        text(s, cx + 0.24, cy + 0.14, cwid - 0.48, 0.24, head, size=11, bold=True, color=INDIGO, char_spacing=1.5)
        for j, (ver, desc) in enumerate(items):
            ry = cy + 0.48 + j * 0.6
            fill, fc = (OLD_TINT, OLD_C) if ver == "V1" else (INDIGO, WHITE)
            box(s, cx + 0.24, ry + 0.11, 0.5, 0.32, fill=fill, radius=0.16)
            text(s, cx + 0.24, ry + 0.11, 0.5, 0.32, ver, size=11, bold=True, color=fc, align="center",
                 valign="middle")
            text(s, cx + 0.88, ry, cwid - 1.1, 0.54, desc, size=12, color=INK, valign="middle", spacing=0.95)

    py = cy + chh + 0.18
    ph = CONTENT_BOTTOM - py
    box(s, 0.5, py, W - 1.0, ph, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    text(s, 0.72, py + 0.14, W - 1.44, 0.28, [[
        ("HOW THE PREPAYMENT SCORE IS PRODUCED", {"bold": True, "color": INDIGO, "char_spacing": 1.5, "size": 11}),
        (f"      Worked example: Lower LLC, May 2026 (#{ex['rank']} client by volume)",
         {"color": MUTE, "size": 11}),
    ]], valign="middle")
    steps = [
        ("Assign each loan to a cohort", "Set once, at lock or funding, from five loan features",
         f"{ex['cohorts']} cohorts", f"{k_fmt(ex['loan_months'])} loan-months"),
        ("Measure the client’s actual CPR", "Annualised rate at which its loan balances run off",
         pct(ex["actual"], 1), None),
        ("Build the look-alike benchmark", "Actual CPR of all loans, from all clients, in those cohorts",
         pct(ex["cohort_bench"], 1), None),
        ("CPRScore = actual ÷ benchmark", "Lower is better: below 1 means slower than look-alike loans",
         f"{ex['cohort_cpr']:.2f}", "Slower than peers" if ex["cohort_cpr"] < 1 else "Faster than peers"),
        ("TotalScore = recapture ÷ CPRScore",
         f"Tagged vs portfolio score: ≥ {FAV_T:g}× Favourable, < {UNF_T:g}× Unfavourable",
         f"{ex['recap']:.2f} ÷ {ex['cohort_cpr']:.2f} = {ex['new']:.2f}", ex["new_band"]),
    ]
    sx0, sgap = 0.72, 0.28
    sw = (W - 1.44 - 4 * sgap) / 5
    sy, sh = py + 0.56, ph - 0.76
    for k, (ttl, desc, val, verdict) in enumerate(steps):
        sx = sx0 + k * (sw + sgap)
        box(s, sx, sy, sw, sh, fill=WHITE, radius=0.07)
        box(s, sx + 0.14, sy + 0.16, 0.36, 0.36, fill=INDIGO, radius=0.18)
        text(s, sx + 0.14, sy + 0.16, 0.36, 0.36, str(k + 1), size=12, bold=True, color=WHITE,
             align="center", valign="middle")
        text(s, sx + 0.6, sy + 0.1, sw - 0.7, 0.5, ttl, size=12.5, bold=True, color=INK, valign="middle",
             spacing=0.92)
        text(s, sx + 0.16, sy + 0.7, sw - 0.32, 0.7, desc, size=11, color=MUTE, spacing=0.95)
        vy = sy + sh - 0.86
        box(s, sx + 0.1, vy, sw - 0.2, 0.74, fill=INDIGO_TINT, radius=0.06)
        vcol = TAG_COL[verdict][0] if verdict in TAG_COL else GREEN if verdict == "Slower than peers" else MUTE
        text(s, sx + 0.16, vy, sw - 0.32, 0.74, [
            [(val, {"size": 17 if len(val) < 14 else 14, "bold": True, "color": INDIGO})],
            *([[(verdict, {"size": 10.5, "bold": True, "color": vcol})]] if verdict else []),
        ], align="center", valign="middle", spacing=0.9)
        if k < 4:
            ax = sx + sw + 0.04
            box(s, ax, sy + sh / 2 - 0.1, sgap - 0.08, 0.2, fill=INDIGO, shape=MSO_SHAPE.RIGHT_ARROW)
    footer(s, page0)

    # =====================================================================
    # COHORT MEMBERSHIP & DEFINITIONS
    # =====================================================================
    s = new_slide(prs)
    tag(s, "COHORTS", INDIGO)
    slide_title(
        s,
        f"{BUILD['cohorts']} look-alike cohorts, assigned once at the pricing decision and fixed "
        "for the life of the loan",
        "Each cohort groups loans that share the features that drive prepayment. A loan’s cohort is "
        "set from its attributes at lock or funding.",
    )

    ky, kh, kg = PY, 0.82, 0.16
    kw = (W - 1.0 - 3 * kg) / 4
    kpis = [
        (str(BUILD["cohorts"]), "cohorts in use", "in the May 2026 scoring run"),
        (k_fmt(loan_months), "loan-months scored", f"{n_clients} clients benchmarked"),
        (pct(mapped_share, 1), "of loan-months mapped", "rest: private-label"),
        (pct(type_share.get("CONV", 0)), "of our balance is Conventional",
         f"FHA {pct(type_share.get('FHA', 0))} · VA {pct(type_share.get('VA', 0))}"),
    ]
    for k, (num, lab, desc) in enumerate(kpis):
        kx = 0.5 + k * (kw + kg)
        box(s, kx, ky, kw, kh, fill=INDIGO_TINT, radius=0.07)
        nw = 0.3 + 0.21 * len(num)
        text(s, kx + 0.2, ky, nw, kh, num, size=26, bold=True, color=INDIGO, valign="middle")
        text(s, kx + 0.25 + nw, ky + 0.06, kw - nw - 0.35, kh - 0.12, [
            [(lab, {"bold": True, "size": 11.5, "color": INK})],
            [(desc, {"size": 9, "color": MUTE})],
        ], valign="middle", spacing=0.95)

    top = ky + kh + 0.16
    ph = 5.95 - top
    DX, DW = 0.5, 6.55
    box(s, DX, top, DW, ph, fill=GRAYFILL, radius=0.08)
    ix, iw = DX + PAD, DW - 2 * PAD
    text(s, ix, top + 0.12, iw, 0.26, [[
        ("DEFINITION", {"bold": True, "color": INDIGO, "char_spacing": 1.5, "size": 10}),
        ("      Five loan features define a cohort", {"bold": True, "size": 12.5, "font": FH}),
    ]], valign="middle")
    lab_w = 1.5
    chip_x0, chip_w = ix + lab_w, iw - lab_w
    dims = [
        ("Loan type", "product", ["Conventional", "FHA", "VA"]),
        ("Loan term", "at origination", ["≤ 20 yr", "> 20 yr"]),
        ("Credit", "original FICO", ["≤ 750", "> 750"]),
        ("Loan size", "balance", ["≤ $150K", "$150–250K", "$250–400K", "> $400K"]),
    ]
    ry0, rstep = top + 0.46, 0.36
    for r, (lab, sub, chips) in enumerate(dims):
        ry = ry0 + r * rstep
        text(s, ix, ry, lab_w - 0.1, 0.34, [[(lab, {"bold": True, "size": 10.5}),
                                             (f"  {sub}", {"size": 8.5, "color": MUTE})]], valign="middle")
        g = 0.06
        cw_ = (chip_w - g * (len(chips) - 1)) / len(chips)
        for j, c in enumerate(chips):
            box(s, chip_x0 + j * (cw_ + g), ry + 0.04, cw_, 0.26, fill=WHITE, line="BEB9F6", line_w=0.75,
                radius=0.13)
            text(s, chip_x0 + j * (cw_ + g), ry + 0.04, cw_, 0.26, c, size=9, bold=True, color=INDIGO,
                 align="center", valign="middle")
    ry = ry0 + 4 * rstep
    text(s, ix, ry, lab_w - 0.1, 0.34, [[("Rate vs market", {"bold": True, "size": 10.5})],
                                        [("note rate − market rate", {"size": 8.5, "color": MUTE})]],
         valign="middle", spacing=0.9)
    nb = len(edges) + 1
    cw_ = chip_w / nb
    zero_at = edges.index(0.0) + 1
    for j in range(nb):
        box(s, chip_x0 + j * cw_ + 0.01, ry + 0.04, cw_ - 0.02, 0.2, fill="CFCBFB" if j < zero_at else "F2D9B0")
    for v, lab in [(0, f"≤ {edges[0]:g}%".replace("-", "−")), (zero_at, "0"), (nb, f"> +{edges[-1]:g}%")]:
        text(s, chip_x0 + v * cw_ - (0.8 if v == nb else 0.4), ry + 0.25, 0.8, 0.16, lab, size=8, color=MUTE, align="center")
    text(s, chip_x0 + chip_w * 0.6, ry + 0.25, chip_w * 0.3, 0.16, "0.25% steps", size=8, color=MUTE,
         align="center")

    my = ry + 0.48
    box(s, ix, my, iw, top + ph - my - 0.12, fill=WHITE, radius=0.06)
    text(s, ix + 0.16, my, iw - 0.3, top + ph - my - 0.12, [
        [("Neighbouring bands that prepay alike are merged → ", {}),
         (f"{BUILD['cohorts']} cohorts in use.", {"bold": True, "color": INDIGO})],
        [(f"Largest, {EX_COHORT} ({pct(ex_cohort_share)} of our balance): ", {"bold": True, "space_before": 3}),
         (describe_rule(EX_RULE) + ".", {})],
    ], size=9.5, color=INK, valign="middle", spacing=0.98)

    MX = DX + DW + 0.2
    MW = W - 0.5 - MX
    box(s, MX, top, MW, ph, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = MX + PAD, MW - 2 * PAD
    text(s, ix, top + 0.12, iw, 0.26, [[
        ("MEMBERSHIP", {"bold": True, "color": INDIGO, "char_spacing": 1.5, "size": 10}),
        ("      Assigned once, at the pricing decision", {"bold": True, "size": 12.5, "font": FH}),
    ]], valign="middle")
    for r, (ch, anchor, note) in enumerate([("Correspondent", "Lock date", "when the rate is locked"),
                                            ("Co-Issue", "Fund date", "when the loan is funded")]):
        ry = top + 0.56 + r * 0.4
        box(s, ix, ry, 1.55, 0.28, fill=WHITE, radius=0.14)
        text(s, ix, ry, 1.55, 0.28, ch, size=9.5, bold=True, color=INK, align="center", valign="middle")
        box(s, ix + 1.63, ry + 0.07, 0.3, 0.14, fill=INDIGO, shape=MSO_SHAPE.RIGHT_ARROW)
        text(s, ix + 2.02, ry, iw - 2.02, 0.28, [[
            (anchor, {"bold": True, "color": INDIGO}), (f"   {note}", {"color": MUTE, "size": 9}),
        ]], size=10.5, valign="middle")
    text(s, ix, top + 1.53, iw, 0.24,
         "Market rate on that date: blend of Freddie Mac PMMS, Optimal Blue and Mortgage News Daily",
         size=8.5, color=MUTE, valign="middle")

    ty = top + 1.86
    text(s, ix, ty, iw, 0.2, [[("Fixed for life: ", {"bold": True, "color": INK}),
                              ("balance and market rates change each month; the cohort does not",
                               {"color": MUTE})]], size=9.5)
    months = ["Lock", "M+1", "M+2", "M+3", "…", "M+36"]
    gw = (iw - 0.1 * (len(months) - 1)) / len(months)
    line(s, ix + gw / 2, ty + 0.5, ix + iw - gw / 2, ty + 0.5, REFGRAY, 1.5)
    for j, m in enumerate(months):
        mx = ix + j * (gw + 0.1)
        if m == "…":
            text(s, mx, ty + 0.36, gw, 0.28, "…", size=12, color=MUTE, align="center", valign="middle")
            continue
        box(s, mx + 0.06, ty + 0.37, gw - 0.12, 0.26, fill=INDIGO if j == 0 else WHITE,
            line=None if j == 0 else INDIGO, line_w=1, radius=0.13)
        text(s, mx + 0.06, ty + 0.37, gw - 0.12, 0.26, EX_COHORT, size=9, bold=True,
             color=WHITE if j == 0 else INDIGO, align="center", valign="middle")
        text(s, mx, ty + 0.65, gw, 0.16, m, size=8, color=MUTE, align="center")

    text(s, ix, ty + 0.9, iw, top + ph - ty - 1.0, [
        [("Definitions can be refreshed: ", {"bold": True}),
         ("if the cohort set is updated, all history is re-scored with the new set, so every month "
          "is compared on the same basis.", {})],
    ], size=9.5, color=INK, valign="middle", spacing=0.98)

    impact_band(s, "WHY FIXED", [
        ("Each client is compared with loans priced like its own. ", {"bold": True, "color": WHITE}),
        ("Because the cohort never moves, a score changes only when prepayment behaviour changes, "
         "not when balances pay down or the rate environment shifts.", {}),
    ], size=11.5)
    footer(s, page0 + 1)
    return 2


def build(path, parts):
    prs = new_deck()
    DECK_MODE["compact"] = True
    page = 1
    for fn in parts:
        page += fn(prs, page)
    prs.save(path)
    print(f"Saved {path}")


if __name__ == "__main__":
    build(OUT, [results_slides, volume_slides, appendix_slides])
    for png_path in render_png(str(OUT), str(ROOT / "B2B_Score_Stability_Exec_Slide_render")):
        print(f"Rendered {png_path}")
    print({"same": same, "up": up, "down": down, "opposite": opposite, "fav": fav_total, "unf": unf_total,
           "share_higher": round(share_higher, 3), "top25_share": round(top25_share, 3), "top25": top25_moves,
           "top25_unf": top25_unf, "total_hits": total_hits, "tag_rate": {k: round(v, 3) for k, v in tag_rate.items()}, "Z": Z, "zero_unf": zero_unf, "rev_zero": rev_zero, "deep": DEEP_DIVE_CLIENTS, "fwd_rev": fwd_rev,
           "hit": hit, "missed": good_flagged_bad, "cohorts": BUILD["cohorts"], "mapped": round(mapped_share, 4),
           "loan_months": loan_months, "type_share": {k: round(v, 3) for k, v in type_share.items()},
           "edges": (edges[0], edges[-1]), "ex_cohort": (EX_COHORT, round(ex_cohort_share, 3), describe_rule(EX_RULE))})
