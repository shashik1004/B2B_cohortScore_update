"""
One-slide path to production + production metrics for the cohort CPRScore release.

Reuses the house-style helpers and back-test numbers from build_stability_slide.py,
so every baseline on this slide matches the 4-slide deck.

    python build_path_to_prod.py   ->  B2B_CPRScore_Path_to_Prod.pptx
"""
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from build_stability_slide import (
    AMBER, AMBER_TINT, GOLD, GRAYFILL, GREEN, GREEN_TINT, INDIGO, INDIGO_TINT, INK, LINEGRAY,
    MUTE, PY, REFGRAY, RESULTS, ROOT, WHITE, FH, H, W,
    agree, box, footer, hit, impact_band, line, n_clients, new_slide, pct, reclass, slide_title,
    tag, tagged, text, total_hits,
)

OUT = ROOT / "B2B_CPRScore_Path_to_Prod.pptx"

# purchase-mix baseline: funded volume by Cohort tag, Sep 2025 – May 2026 fundings
detail = pd.read_excel(RESULTS, sheet_name="Base Approach Detail")
mix = detail.groupby("CohortBand")["CohortFundedLoanAmt"].sum()
funded_total = mix.sum()
mix_share = (mix / funded_total).to_dict()

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
s = new_slide(prs)
tag(s, "ROLLOUT", INDIGO)
slide_title(
    s,
    "With a go decision on 10 Oct, the cohort CPRScore can be live by end of October, "
    "tracked on three production metrics",
    "This release changes the CPR part of the client score only; the recapture score stays as it is "
    "today and is refreshed in the next phase.",
)

# ----- roadmap -----
ry, rh = PY, 1.5
box(s, 0.5, ry, W - 1.0, rh, fill=GRAYFILL, radius=0.08)
steps = [
    ("DONE", GREEN, True, "Sep 2026", "Methodology & back-test", "Cohorts built and back-tested"),
    ("DECISION", AMBER, True, "10 Oct", "Go / No-Go", "B2B team decision"),
    ("GO-LIVE", INDIGO, True, "By 31 Oct", "Cohort CPRScore in production", "Recapture score unchanged"),
    ("TRACK", INDIGO, False, "From Nov", "Production metrics, monthly", "First 3-month read ~3 months after launch"),
    ("NEXT PHASE", MUTE, False, "After go-live", "Recapture score refresh", "Picked up as the next phase"),
]
n = len(steps)
x0, x1 = 0.72, W - 0.72
colw = (x1 - x0) / n
node_y = ry + 0.52
line(s, x0 + colw / 2, node_y, x1 - colw / 2, node_y, REFGRAY, 2)
for k, (status, col, filled, date, ttl, desc) in enumerate(steps):
    cx = x0 + k * colw
    mid = cx + colw / 2
    d = 0.22
    box(s, mid - d / 2, node_y - d / 2, d, d, fill=col if filled else WHITE, line=None if filled else col,
        line_w=2, shape=MSO_SHAPE.OVAL)
    cw = 1.18
    box(s, mid - cw / 2, ry + 0.14, cw, 0.24, fill=col if filled else WHITE, line=None if filled else col,
        line_w=1, radius=0.12)
    text(s, mid - cw / 2, ry + 0.14, cw, 0.24, status, size=8.5, bold=True,
         color=WHITE if filled else col, align="center", valign="middle", char_spacing=1)
    text(s, cx + 0.08, node_y + 0.17, colw - 0.16, 0.78, [
        [(date, {"size": 14, "bold": True, "color": col if col != MUTE else INK})],
        [(ttl, {"size": 10.5, "bold": True, "color": INK})],
        [(desc, {"size": 9, "color": MUTE})],
    ], align="center", spacing=0.95)

# ----- production metrics -----
my = ry + rh + 0.14
text(s, 0.5, my, W - 1.0, 0.26, [[
    ("PRODUCTION METRICS TO TRACK", {"bold": True, "color": INDIGO, "char_spacing": 1.5, "size": 10}),
    ("      Baselines come from the back-test, so each month can be compared with what we expected",
     {"color": MUTE, "size": 10}),
]], valign="middle")
cy = my + 0.32
ch = 5.95 - cy
cg = 0.2
cw = (W - 1.0 - 2 * cg) / 3
fav, unf = "Favourable", "Unfavourable"
cards = [
    ("Prediction", "Do clients with a high CPRScore actually prepay faster than their peers 3 months "
     "later, and low scorers slower?",
     "Share of Favourable and Unfavourable tags confirmed by the next 3 months’ performance",
     [(pct(hit[("Cohort", fav)] / tagged[("Cohort", fav)]), "Fav. confirmed"),
      (pct(hit[("Cohort", unf)] / tagged[("Cohort", unf)]), "Unfav. confirmed")],
     f"{total_hits['Cohort']} correct calls of {n_clients} clients in the back-test",
     "A sustained drop below back-test levels"),
    ("Stability", "Do client tags hold as each month’s new data comes in, or flip back and forth?",
     "Share of clients keeping their tag from one refresh to the next; number changing tag",
     [(pct(agree["Cohort"]), "keep their tag"), (str(reclass["Cohort"]), "clients change tag")],
     "May scores re-run with three more months of data",
     "Spikes in tag changes among large clients"),
    ("Purchase mix", "How much are we buying from Favourable vs Unfavourable clients, and how is the "
     "mix changing?",
     "Funded volume ($ and loans) by client tag, each month",
     None,
     f"${funded_total / 1e9:.1f}B funded, Sep 2025 – May 2026",
     "Favourable share rising over time as pricing responds"),
]
for k, (name, q, measure, base, base_note, watch) in enumerate(cards):
    cx = 0.5 + k * (cw + cg)
    box(s, cx, cy, cw, ch, fill=INDIGO_TINT if k == 0 else GRAYFILL,
        line=INDIGO if k == 0 else None, line_w=1.25, radius=0.08)
    ix, iw = cx + 0.2, cw - 0.4
    box(s, ix, cy + 0.16, 0.3, 0.3, fill=INDIGO, radius=0.15)
    text(s, ix, cy + 0.16, 0.3, 0.3, str(k + 1), size=10.5, bold=True, color=WHITE, align="center",
         valign="middle")
    text(s, ix + 0.42, cy + 0.14, iw - 0.42, 0.34, name, size=14, bold=True, font=FH, valign="middle")
    text(s, ix, cy + 0.52, iw, 0.4, q, size=9.5, color=INK, spacing=0.95)
    text(s, ix, cy + 0.94, iw, 0.3, [[("Measure  ", {"bold": True, "color": INDIGO, "size": 8.5,
                                                      "char_spacing": 1}), (measure, {})]],
         size=9, color=MUTE, spacing=0.95)

    by, bh = cy + 1.28, 0.62
    box(s, ix, by, iw, bh, fill=WHITE, radius=0.06)
    text(s, ix + 0.12, by + 0.06, iw - 0.24, 0.18, "BACK-TEST BASELINE", size=8, bold=True, color=MUTE,
         char_spacing=1)
    if base:
        half = (iw - 0.24) / 2
        for j, (num, lab) in enumerate(base):
            text(s, ix + 0.12 + j * half, by + 0.2, half, 0.26, [[
                (num, {"size": 19, "bold": True, "color": INDIGO}),
                (f"  {lab}", {"size": 8.5, "color": INK}),
            ]], valign="middle")
    else:
        bx, bw = ix + 0.12, iw - 0.24
        sx = bx
        for band, col, fc in [(fav, GREEN, WHITE), ("Neutral", REFGRAY, INK), (unf, AMBER, WHITE)]:
            w_ = bw * mix_share[band]
            box(s, sx, by + 0.22, w_, 0.22, fill=col)
            text(s, sx, by + 0.22, w_, 0.22, f"{band[:3] if band != 'Neutral' else 'Neu'}. "
                 f"{pct(mix_share[band])}", size=8.5, bold=True, color=fc, align="center", valign="middle")
            sx += w_
    text(s, ix + 0.12, by + 0.45, iw - 0.24, 0.16, base_note, size=8, color=MUTE)

    wh = ch - (by - cy) - bh - 0.1
    assert wh > 0.15, wh
    text(s, ix, by + bh + 0.04, iw, wh, [[
        ("Watch for  ", {"bold": True, "color": AMBER, "size": 8.5, "char_spacing": 1}), (watch, {}),
    ]], size=9, color=INK, valign="middle", spacing=0.95)

impact_band(s, "DECISION NEEDED", [
    ("Go / No-Go from the B2B team on 10 Oct. ", {"bold": True, "color": WHITE}),
    ("On a go, the cohort CPRScore goes live by ", {}), ("end of October", {"bold": True, "color": GOLD}),
    (" with these three metrics reported monthly; the recapture score refresh follows as the next phase.", {}),
], size=11.5)
footer(s, 1)
prs.save(OUT)
print(f"Saved {OUT}")
print({k: round(v, 3) for k, v in mix_share.items()}, round(funded_total / 1e9, 2))
