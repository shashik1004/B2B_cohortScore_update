"""
One-page MBR summary of the cohort-benchmark proposal (exec compression of B2B_CPRScore_Exec_Deck.pptx).

Reuses the numbers and house-style helpers from build_stability_slide.py; does not rebuild the deck.

    python build_mbr_slide.py   ->  B2B_CPRScore_MBR_Slide.pptx
"""
from pptx import Presentation
from pptx.util import Inches

import build_stability_slide as b
from build_stability_slide import (
    AMBER, BANDS, FAV_T, FH, GREEN, GRAYFILL, INDIGO, INDIGO_DK, INDIGO_TEXT_LT, INDIGO_TEXT_MUTE, INDIGO_TINT,
    INK, MUTE, OLD_C, PAD, RED, REFGRAY, UNF_T, W, WHITE, H, box, footer, line, new_slide, panel_header, pct,
    slide_title, tag, text,
)

OUT = b.ROOT / "B2B_CPRScore_MBR_Slide.pptx"
TOP, BOT = 1.75, 5.45            # panel band
ASK_Y, ASK_H = 5.6, 1.27         # sign-off band, ends just above the footer bar


def build():
    b.DECK_MODE["compact"] = False                     # laid out directly; no stretch
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    s = new_slide(prs)
    tag(s, "MBR · PROPOSAL", INDIGO)
    n = b.n_clients
    slide_title(
        s,
        f"Proposal: score client prepayment against actual look-alike loans ({b.BUILD['cohorts']} cohorts); "
        "tags hold up better, and we are seeking B2B sign-off",
        f"May 2026 TotalScore · {n} clients · {b.bn(b.vol_tot)} 12-month funded volume · each approach tagged "
        f"against its own portfolio score (Favourable ≥ {FAV_T:g}×, Unfavourable < {UNF_T:g}×)",
    )
    ph = BOT - TOP

    # ----- 1. The proposal -----
    x1, w1 = 0.5, 3.7
    box(s, x1, TOP, w1, ph, fill=GRAYFILL, radius=0.08)
    ix, iw = x1 + PAD, w1 - 2 * PAD
    panel_header(s, ix, TOP + 0.14, iw, 1, "THE PROPOSAL", "Measure prepayment against what look-alike loans did")
    rows = [("OLD", OLD_C, "Client CPR ÷ model-predicted CPR for its loans"),
            ("COHORT", INDIGO, f"Client CPR ÷ actual CPR of every lender’s loans in the same {b.BUILD['cohorts']} "
                               "cohorts")]
    for k, (lab, col, desc) in enumerate(rows):
        yy = TOP + 1.0 + k * 0.56
        box(s, ix, yy, iw, 0.5, fill=WHITE if lab == "OLD" else INDIGO_TINT, radius=0.05,
            line=INDIGO if lab == "COHORT" else None, line_w=1.0)
        text(s, ix + 0.1, yy, 0.72, 0.5, lab, size=8.5, bold=True, color=col, valign="middle", char_spacing=1)
        text(s, ix + 0.82, yy, iw - 0.9, 0.5, desc, size=9, color=INK, valign="middle", spacing=0.92)
    model_cpr, actual_cpr = float(b.old_total["ModelCPR"]), float(b.old_total["ActualCPR"])
    text(s, ix, TOP + 2.16, iw, 0.62, [
        [("Why: ", {"bold": True}), (f"the model expects {pct(model_cpr, 1)} prepayment across the book against "
                                     f"{pct(actual_cpr, 1)} actual, and misses by a different amount for each loan "
                                     "type, which reorders clients.", {})],
        [("Cohorts are set once, at lock or funding, from loan type, FICO, term, size and rate vs market.",
          {"color": MUTE, "space_before": 4})],
    ], size=8.5, color=INK, spacing=0.93)
    sy = TOP + 2.95
    line(s, ix, sy - 0.06, ix + iw, sy - 0.06, REFGRAY, 0.75)
    text(s, ix, sy, iw, 0.24, "SAME TAG UNDER BOTH APPROACHES", size=7.5, bold=True, color=MUTE, char_spacing=1)
    for k, (v, lab) in enumerate([(b.same / n, "of clients"), (b.vsame, "of volume")]):
        text(s, ix + k * iw / 2, sy + 0.2, iw / 2, 0.5, [[(pct(v), {"size": 22, "bold": True, "color": INDIGO,
                                                                     "font": FH})], [(lab, {"size": 8, "color": MUTE})]],
             spacing=0.85)

    # ----- 2. Stability and misclassification (emphasis) -----
    x2, w2 = x1 + w1 + 0.2, 4.25
    box(s, x2, TOP, w2, ph, fill=INDIGO_TINT, line=INDIGO, line_w=1.25, radius=0.08)
    ix, iw = x2 + PAD, w2 - 2 * PAD
    panel_header(s, ix, TOP + 0.14, iw, 2, "STABILITY & MISCLASSIFICATION",
                 "Cohort tags hold up better and mis-flag fewer good clients")
    text(s, ix, TOP + 0.94, iw, 0.5, [[
        (pct(b.tag_rate["Cohort"], 1), {"size": 28, "bold": True, "color": INDIGO, "font": FH}),
        ("  vs  ", {"size": 11, "color": MUTE}),
        (pct(b.tag_rate["Old"], 1), {"size": 28, "bold": True, "color": OLD_C, "font": FH}),
    ]], valign="middle")
    text(s, ix, TOP + 1.44, iw, 0.2, f"of Fav. / Unfav. tags confirmed four months later ({b.F_BASE} → {b.F_COMP})",
         size=8.5, color=MUTE)
    cw = 0.7
    c0 = iw - 2 * cw
    ty = TOP + 1.7
    for lab, col, x_ in [("Cohort", INDIGO, ix + c0), ("Old", OLD_C, ix + c0 + cw)]:
        text(s, x_, ty, cw, 0.18, lab, size=8, bold=True, color=col, align="center")
    miss = {a: b.good_flagged_bad[a] / b.tagged[(a, "Unfavourable")] for a in b.APPROACHES}
    trows = [
        ("Unfavourable tags confirmed", b.hit_rate[("Cohort", "Unfavourable")], b.hit_rate[("Old", "Unfavourable")], True),
        ("Favourable tags confirmed", b.hit_rate[("Cohort", "Favourable")], b.hit_rate[("Old", "Favourable")], True),
        ("Good clients tagged Unfavourable (clients)", miss["Cohort"], miss["Old"], False),
        ("Good clients tagged Unfavourable (volume)", b.v_missed["Cohort"], b.v_missed["Old"], False),
    ]
    rh = 0.27
    for k, (lab, cv, ov, hi_good) in enumerate(trows):
        yy = ty + 0.22 + k * rh
        box(s, ix - 0.05, yy, iw + 0.1, rh - 0.03, fill=WHITE, radius=0.04)
        text(s, ix + 0.04, yy, c0 - 0.04, rh - 0.03, lab, size=8.5, color=INK, valign="middle")
        better_c = cv > ov if hi_good else cv < ov
        for x_, v, col, bold_ in [(ix + c0, cv, INDIGO, better_c), (ix + c0 + cw, ov, OLD_C, not better_c)]:
            text(s, x_, yy, cw, rh - 0.03, pct(v, 1) if hi_good else pct(v), size=10, bold=bold_, color=col,
                 align="center", valign="middle")
    ny = ty + 0.22 + len(trows) * rh + 0.06
    text(s, ix, ny, iw, BOT - 0.06 - ny, [
        [("Good clients tagged Unfavourable = Unfavourable tags that turned Favourable later (lower is better). "
          "By volume, overall confirmation is level ", {}),
         (f"({pct(b.v_tag_rate['Cohort'])} vs {pct(b.v_tag_rate['Old'])})", {"bold": True}),
         ("; the edge is on Unfavourable tags.", {})],
    ], size=7.5, color=MUTE, spacing=0.92)

    # ----- 3. Client-level implications -----
    x3 = x2 + w2 + 0.2
    w3 = W - 0.5 - x3
    box(s, x3, TOP, w3, ph, fill=GRAYFILL, radius=0.08)
    ix, iw = x3 + PAD, w3 - 2 * PAD
    panel_header(s, ix, TOP + 0.14, iw, 3, "CLIENT-LEVEL IMPLICATIONS",
                 "Cohort is the stricter read: more volume tagged Unfavourable")
    text(s, ix, TOP + 1.04, iw, 0.18, [[("Share of funded volume   ", {"bold": True, "color": INK}),
                                        ("■ ", {"color": GREEN}), ("Fav.  ", {}), ("■ ", {"color": REFGRAY}),
                                        ("Neutral  ", {}), ("■ ", {"color": RED}), ("Unfav.", {})]], size=8, color=MUTE)
    lw = 0.62
    for k, (a, col) in enumerate([("Old", OLD_C), ("Cohort", INDIGO)]):
        ry = TOP + 1.26 + k * 0.32
        text(s, ix, ry, lw - 0.06, 0.26, a, size=9, bold=True, color=col, valign="middle")
        sx = ix + lw
        for bd, fill, fc in [("Favourable", GREEN, WHITE), ("Neutral", REFGRAY, INK), ("Unfavourable", RED, WHITE)]:
            sw = (iw - lw) * b.vshare[a][bd]
            box(s, sx, ry, sw, 0.26, fill=fill)
            text(s, sx, ry, sw, 0.26, pct(b.vshare[a][bd]), size=8.5, bold=True, color=fc, align="center",
                 valign="middle")
            sx += sw
    apm, tomo = b.DIVES
    fav_need = b.need[("Cohort", "Favourable")], b.need[("Old", "Favourable")]
    bullets = [
        [(f"{b.down} clients tag lower and {b.up} higher; ", {"bold": True}),
         (f"{b.unf_total['Cohort']} vs {b.unf_total['Old']} Unfavourable. Of the top 25 lenders, "
          f"{b.top25_moves['same']} keep their tag, {b.top25_moves['down']} move lower.", {})],
        [("Deep dives: ", {"bold": True}),
         (f"{b.short_name(apm['name'])} (#{apm['rank']}) prepays {pct(apm['cpr_score'] - 1)} faster than look-alike "
          f"loans ({pct(apm['client'], 1)} vs {pct(apm['bench'], 1)}); {b.short_name(tomo['name'])} (#{tomo['rank']}) "
          f"{pct(tomo['cpr_score'] - 1)} faster, almost all in jumbo loans. Both tag Unfavourable.", {})],
        [("Buying implication: ", {"bold": True}),
         (f"an 85 / 10 / 5 buy mix at today’s ~{pct(b.WIN_TODAY)} win rate needs {pct(fav_need[0], 1)} on "
          f"Favourable volume ({pct(fav_need[1], 1)} on Old tags).", {})],
    ]
    by = TOP + 2.0
    text(s, ix, by, iw, BOT - 0.08 - by,
         [[("▪ ", {"color": INDIGO, "bold": True, **({"space_before": 5} if k else {})}), *p] for k, p in enumerate(bullets)],
         size=8.5, color=INK, spacing=0.93)

    # ----- Sign-off band -----
    box(s, 0.5, ASK_Y, W - 1.0, ASK_H, fill=INDIGO_DK, radius=0.08)
    text(s, 0.75, ASK_Y + 0.12, 3.0, 0.22, "SIGN-OFF REQUESTED FROM B2B", size=10, bold=True, color=INDIGO_TEXT_MUTE,
         char_spacing=1.5)
    asks = [
        ("Adopt the cohort benchmark", "Replace model CPR with look-alike actual CPR in the prepayment score"),
        ("Approve go-live", "Generate the new Cohort scores from the first week of Nov / Dec"),
        ("Monitor monthly / quarterly", "Track prediction stability and purchase mix month over month"),
        (f"Target ~{pct(fav_need[0], 1)} win rate on Favourable",
         f"Up from ~{pct(b.WIN_TODAY)} today, to buy 85 / 10 / 5 at today’s volume ({pct(fav_need[1], 1)} on Old tags)"),
    ]
    aw = (W - 1.0 - 0.5 - 3 * 0.2) / 4
    for k, (head, desc) in enumerate(asks):
        ax = 0.75 + k * (aw + 0.2)
        ay = ASK_Y + 0.42
        box(s, ax, ay + 0.02, 0.28, 0.28, fill=INDIGO, radius=0.14)
        text(s, ax, ay + 0.02, 0.28, 0.28, str(k + 1), size=10, bold=True, color=WHITE, align="center",
             valign="middle")
        text(s, ax + 0.38, ay, aw - 0.4, 0.78, [[(head, {"bold": True, "color": WHITE, "size": 10.5})],
                                                [(desc, {"color": INDIGO_TEXT_LT, "size": 8.5})]],
             spacing=0.92)
    footer(s, 1)
    prs.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    build()
