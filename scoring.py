"""
scoring.py — Fragment scoring
==============================
Score = sum(weight * roll_quality) / ideal * 100

roll_quality = value / (max_roll * avg_rolls_per_stat)
  where avg_rolls_per_stat = (upgrade+1) / 4

This normalizes against what a "well-rolled" fragment would have,
not against the near-impossible dream max.

At +5: avg_rolls_per_stat = 5/4 = 1.25
A stat at exactly average roll value × 1.25 rolls = roll_quality 1.0
A stat with high rolls or multiple rolls = roll_quality > 1.0 (capped at ~2.0)

Grades:
  SSS >= 85   SS >= 70   S >= 55
  A   >= 40   B  >= 28   C >= 16   D >= 7   F < 7
"""

from data import CHARS, SET_TO_CHARS, FIXED_MAIN, ROLL_RANGES

_MAX_ROLL = {k: v[1] for k, v in ROLL_RANGES.items()}
_AVG_ROLL = {k: (v[0]+v[1])/2 for k, v in ROLL_RANGES.items()}

GRADES = [
    (88, "SSS", "#d966ff", "Keep"),
    (72, "SS",  "#FFD700", "Keep"),
    (57, "S",   "#FFB347", "Keep"),
    (43, "A",   "#58a6ff", "Keep"),
    (30, "B",   "#7ec87e", "Maybe"),
    (18, "C",   "#EF9F27", "Discard"),
    ( 8, "D",   "#c87e7e", "Discard"),
    ( 0, "F",   "#484848", "Discard"),
]

def _grade(score: float):
    for threshold, grade, color, verdict in GRADES:
        if score >= threshold:
            return grade, color, verdict
    return "F", "#484848", "Discard"


def _roll_quality(stat: str, value: float, upgrade: int) -> float:
    """How many effective max-rolls worth of value does this stat have?
    1.0 = one perfect roll. 2.5 = two and a half perfect rolls (excellent).
    Capped at 5.0 (theoretical max = all 5 rolls on one stat at max value)."""
    if stat not in _MAX_ROLL or _MAX_ROLL[stat] <= 0:
        return 0.0
    return min(5.0, value / _MAX_ROLL[stat])


def score_for_char(char_name: str, subs: list, main_name: str,
                   slot: str, upgrade: int = 1) -> dict:
    cd       = CHARS[char_name]
    cw       = cd["weights"]
    is_maxed = (upgrade >= 5)

    # Ideal: top-4 weights × 2.5
    # (represents a fragment where each good stat received ~2.5 effective rolls at max value)
    top4_w = sorted(cw.values(), reverse=True)[:4]
    ideal  = sum(top4_w) * 2.5 or 1.0

    # ── Current score ──────────────────────────────────────────────────────────
    sub_score = 0.0
    details   = []
    for s in subs:
        w  = cw.get(s["name"], 0.0)
        rq = _roll_quality(s["name"], s["fval"], upgrade)
        contrib = w * rq
        sub_score += contrib
        details.append({
            "name":    s["name"],
            "value":   s["value"],
            "w":       w,
            "rv":      round(rq, 3),
            "contrib": round(contrib, 3),
        })

    current_pct = (sub_score / ideal) * 100.0

    # ── Best at +5 ─────────────────────────────────────────────────────────────
    if is_maxed:
        best_pct = current_pct
    else:
        remaining  = 5 - upgrade
        good_subs  = sorted(subs, key=lambda s: cw.get(s["name"], 0.0), reverse=True)
        best_score = sub_score
        if good_subs:
            best_stat  = good_subs[0]["name"]
            best_w     = cw.get(best_stat, 0.0)
            cur_rq     = _roll_quality(best_stat, good_subs[0]["fval"], upgrade)
            # Future rolls all go to best stat at max roll value
            max_r      = _MAX_ROLL.get(best_stat, 0)
            future_val = good_subs[0]["fval"] + remaining * max_r
            new_rq     = _roll_quality(best_stat, future_val, 5)
            best_score = sub_score - best_w * cur_rq + best_w * new_rq
        best_pct = (best_score / ideal) * 100.0

    # ── Main stat ──────────────────────────────────────────────────────────────
    main_ok    = None
    main_bonus = 0.0
    if slot and main_name:
        fixed = FIXED_MAIN.get(slot)
        if fixed:
            main_ok = (main_name == fixed)
            if not main_ok: main_bonus = -8.0
        else:
            gm = cd.get("good_main", {}).get(slot)
            if gm:
                gl        = gm if isinstance(gm, list) else [gm]
                main_ok   = main_name in gl
                main_bonus = +5.0 if main_ok else -10.0

    score     = round(max(0.0, min(100.0, current_pct + main_bonus)), 1)
    max_score = round(max(0.0, min(100.0, best_pct    + main_bonus)), 1)

    grade,      color,      verdict      = _grade(score if is_maxed else max_score)
    best_grade, best_color, best_verdict = _grade(max_score)

    return {
        "char":        char_name,
        "score":       score,
        "max_score":   max_score,
        "grade":       grade,
        "best_grade":  best_grade,
        "color":       color,
        "best_color":  best_color,
        "verdict":     verdict,
        "details":     details,
        "main_ok":     main_ok,
        "main_bonus":  main_bonus,
        "is_maxed":    is_maxed,
        "n_good_subs": sum(1 for s in subs if cw.get(s["name"], 0.0) >= 0.5),
        "recommended": True,
        "wr":          round(sub_score, 2),
        "best_wr":     round(best_score if not is_maxed else sub_score, 2),
    }


def score_fragment(frag: dict) -> dict:
    set_key   = frag.get("set")
    slot      = frag.get("slot")
    subs      = frag.get("sub_stats", [])
    main      = frag.get("main_stat")
    main_name = main["name"] if main else None
    try:
        upgrade = int(str(frag.get("upgrade","1")).replace("+",""))
    except Exception:
        upgrade = 1

    _empty = {
        "score":0,"max_score":0,"grade":"?","best_grade":"?",
        "verdict":"Unknown","color":"#7d8590","details":[],
        "main_ok":None,"chars":[],"set_found":False,
        "best_char":None,"char_scores":[],"upgrade":upgrade,
        "is_maxed":upgrade>=5,"wr":0,"best_wr":0,
    }
    if not set_key:
        return {**_empty, "verdict":"Set not recognised"}

    chars_for_set = SET_TO_CHARS.get(set_key, list(CHARS.keys()))
    all_chars     = list(CHARS.keys())

    char_scores = []
    for ch in all_chars:
        cs = score_for_char(ch, subs, main_name, slot, upgrade)
        cs["recommended"] = ch in chars_for_set
        char_scores.append(cs)

    char_scores.sort(key=lambda x: (
        0 if x["recommended"] else 1,
        -(x["max_score"] if not x["is_maxed"] else x["score"])
    ))

    best = char_scores[0]
    grade_score = best["score"] if best["is_maxed"] else best["max_score"]
    grade, color, verdict = _grade(grade_score)

    return {
        "score":       best["score"],
        "max_score":   best["max_score"],
        "grade":       grade,
        "best_grade":  best["best_grade"],
        "verdict":     verdict,
        "color":       color,
        "details":     best["details"],
        "main_ok":     best["main_ok"],
        "chars":       chars_for_set,
        "set_found":   True,
        "best_char":   best["char"],
        "char_scores": char_scores,
        "upgrade":     upgrade,
        "n_good_subs": best["n_good_subs"],
        "is_maxed":    best["is_maxed"],
        "wr":          best["wr"],
        "best_wr":     best["best_wr"],
    }
