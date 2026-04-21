"""
CZN Memory Fragment Overlay v4
================================
Windows OCR (WinRT via PowerShell) — no Tesseract needed
Scoring: Fribbels-style roll value per character
Hotkey: F9 (global low-level hook)
"""

import tkinter as tk
import threading, re, os, sys, ctypes, ctypes.wintypes as wt

try:
    import pyautogui
    from PIL import ImageGrab, Image
except ImportError as e:
    print(f"Missing library: {e}\npip install pyautogui Pillow")
    sys.exit(1)

# ── OCR (Tesseract) ──────────────────────────────────────────────────────────
import pytesseract
from PIL import ImageFilter, ImageEnhance

for _p in [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    rf"C:\Users\{os.environ.get('USERNAME','')}\AppData\Local\Tesseract-OCR\tesseract.exe",
]:
    if os.path.exists(_p):
        pytesseract.pytesseract.tesseract_cmd = _p
        break

def run_ocr(img: Image.Image) -> str:
    """Multi-pass OCR. First pass determines line order (top-to-bottom).
    Additional passes only add lines that were missed by the first pass.
    This preserves the top-to-bottom order needed for correct Main-Stat detection.
    """
    w, h = img.size
    img = img.resize((w * 3, h * 3), Image.LANCZOS)

    def ocr_pass(contrast, threshold, psm):
        p = img.copy().convert("L")
        p = ImageEnhance.Contrast(p).enhance(contrast)
        p = p.point(lambda px: 0 if px < threshold else 255, "L")
        p = p.filter(ImageFilter.SHARPEN)
        try:
            return pytesseract.image_to_string(p, config=f"--psm {psm} -l eng")
        except Exception:
            return ""

    # Pass 1 is the authoritative pass — sets the line order
    primary = ocr_pass(3.0, 160, 6)
    primary_lines = [l.strip() for l in primary.splitlines() if l.strip()]

    # Additional passes — only used to recover lines missed by pass 1
    extra_passes = [
        ocr_pass(3.0, 130, 6),   # lower threshold — catches grey sub-stats
        ocr_pass(4.0, 180, 6),   # higher contrast — catches faint text
        ocr_pass(2.5, 140, 11),  # sparse mode — isolated lines
    ]

    # Build dedup set from primary pass
    seen = {re.sub(r'\s+', ' ', l.lower()) for l in primary_lines}

    # Append extra lines found in later passes (added at end, not interleaved)
    extra_lines = []
    for text in extra_passes:
        for line in text.splitlines():
            line = line.strip()
            if not line or len(line) < 3:
                continue
            key = re.sub(r'\s+', ' ', line.lower())
            if key not in seen:
                seen.add(key)
                extra_lines.append(line)

    return "\n".join(primary_lines + extra_lines)

# ── Low-Level Keyboard Hook ───────────────────────────────────────────────────
WH_KEYBOARD_LL = 13
WM_KEYDOWN     = 0x0100
VK_F9          = 0x78
VK_ESCAPE      = 0x1B
HOOKPROC = ctypes.CFUNCTYPE(ctypes.c_long, ctypes.c_int, wt.WPARAM, wt.LPARAM)

class LowLevelHook:
    def __init__(self):
        self._hook = None; self._cb = None
        self._cbs = {}  # dict[int, list]

    def add(self, vk: int, fn):
        self._cbs.setdefault(vk, []).append(fn)

    def start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        u = ctypes.windll.user32
        def _proc(nCode, wParam, lParam):
            if nCode >= 0 and wParam == WM_KEYDOWN:
                try:
                    # lParam points to KBDLLHOOKSTRUCT; first field is vkCode (DWORD)
                    vk = ctypes.cast(lParam, ctypes.POINTER(ctypes.c_uint32))[0]
                    for fn in self._cbs.get(vk, []):
                        threading.Thread(target=fn, daemon=True).start()
                except Exception:
                    pass
            try:
                return u.CallNextHookEx(self._hook, nCode, wParam, lParam)
            except Exception:
                return 0
        self._cb = HOOKPROC(_proc)
        self._hook = u.SetWindowsHookExW(WH_KEYBOARD_LL, self._cb, None, 0)
        msg = wt.MSG()
        while u.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            u.TranslateMessage(ctypes.byref(msg))
            u.DispatchMessageW(ctypes.byref(msg))

    def stop(self):
        if self._hook:
            ctypes.windll.user32.UnhookWindowsHookEx(self._hook)

_hook = LowLevelHook()

# ── Colours ───────────────────────────────────────────────────────────────────
C = {
    "bg":"#0d1117","border":"#21262d",
    "p1_bg":"#003A8C","p1_fg":"#FFD700",
    "p2_bg":"#FFD700","p2_fg":"#003A8C",
    "bad_bg":"#1c1c1c","bad_fg":"#484848",
    "main_ok_bg":"#0d2d0d","main_ok_fg":"#7ec87e",
    "main_bad_bg":"#2d0d0d","main_bad_fg":"#c87e7e",
    "main_neu_bg":"#1a1a2a","main_neu_fg":"#8888bb",
    "text":"#e6edf3","muted":"#7d8590",
    "s":"#FFD700","a":"#58a6ff","b":"#7ec87e","c":"#EF9F27","d":"#c87e7e","f":"#484848",
}

# ── Slot system ───────────────────────────────────────────────────────────────
# Shock=I, Suppression=II, Denial=III → fixed main stats
# Ideal=IV, Desire=V, Imagination=VI → variable main stats
FIXED_MAIN = {"shock":"ATK","suppression":"DEF","denial":"HP"}
SLOT_NAMES = {
    "shock":"shock","i":"shock","1":"shock",
    "suppression":"suppression","ii":"suppression","2":"suppression",
    "denial":"denial","iii":"denial","3":"denial",
    "ideal":"ideal","iv":"ideal","4":"ideal",
    "desire":"desire","v":"desire","5":"desire",
    "imagination":"imagination","vi":"imagination","6":"imagination",
}
SLOT_DISPLAY = {
    "shock":"I · Shock","suppression":"II · Suppression","denial":"III · Denial",
    "ideal":"IV · Ideal","desire":"V · Desire","imagination":"VI · Imagination",
}

# ── Roll ranges (Legendary +5) ────────────────────────────────────────────────
# (min_per_roll, max_per_roll) for each sub-stat on Legendary fragments
# Source: Prydwen stats guide + community data
ROLL_RANGES = {
    "Crit Rate%":   (1.6, 3.2),
    "Crit DMG%":    (3.2, 6.4),
    "DEF":          (3.0, 6.0),
    "DEF%":         (1.6, 3.2),
    "ATK":          (5.0, 9.0),
    "ATK%":         (1.6, 3.2),
    "HP":           (8.0,16.0),
    "HP%":          (1.6, 3.2),
    "EGO Recovery": (4.0, 8.0),
    "DoT%":         (1.6, 3.2),
    "Extra DMG%":   (1.6, 3.2),
}
# Max value a sub-stat can reach at +5 (4 rolls total, each at max)
MAX_VALS = {k: v[1]*4 for k,v in ROLL_RANGES.items()}
MIN_VALS = {k: v[0]*1 for k,v in ROLL_RANGES.items()}  # baseline (1 min roll)

def roll_value(stat: str, value: float) -> float:
    """0.0-1.0: how good is this roll relative to perfect 4x max rolls."""
    if stat not in ROLL_RANGES:
        return 0.0
    mn = MIN_VALS[stat]
    mx = MAX_VALS[stat]
    if mx <= mn:
        return 0.0
    return max(0.0, min(1.0, (value - mn) / (mx - mn)))

# ── Stat weights per character/role ──────────────────────────────────────────
# 0.0 = useless, 0.5 = ok, 0.75 = good, 1.0 = best
# Inspired by Fribbels methodology; calibrated against Prydwen substat priority lists

CRIT_DPS = {
    "Crit Rate%":1.0,"Crit DMG%":1.0,"ATK":0.75,"ATK%":0.5,
    "DEF":0.0,"DEF%":0.0,"HP":0.0,"HP%":0.0,"EGO Recovery":0.0,
    "DoT%":0.0,"Extra DMG%":0.0,
}
NINE_W = {  # Nine needs DEF for Potential 7 (350 DEF goal)
    "Crit Rate%":1.0,"Crit DMG%":1.0,"DEF":0.75,"DEF%":0.5,
    "ATK":0.25,"ATK%":0.1,"HP":0.0,"HP%":0.0,"EGO Recovery":0.0,
    "DoT%":0.0,"Extra DMG%":0.0,
}
TIPHERA_W = {  # DEF scaler
    "DEF":1.0,"DEF%":0.9,"Crit Rate%":0.6,"Crit DMG%":0.6,
    "HP":0.1,"HP%":0.1,"ATK":0.0,"ATK%":0.0,"EGO Recovery":0.1,
    "DoT%":0.0,"Extra DMG%":0.0,
}
KHALIPE_W = {  # needs both crit AND DEF
    "Crit Rate%":1.0,"Crit DMG%":0.9,"DEF":0.9,"DEF%":0.7,
    "EGO Recovery":0.4,"ATK":0.1,"ATK%":0.1,"HP":0.0,"HP%":0.0,
    "DoT%":0.0,"Extra DMG%":0.0,
}
DEF_SUP_W = {  # Orlea, Cassius support
    "DEF":1.0,"DEF%":0.9,"HP":0.6,"HP%":0.5,"EGO Recovery":0.5,
    "Crit Rate%":0.0,"Crit DMG%":0.0,"ATK":0.0,"ATK%":0.0,
    "DoT%":0.0,"Extra DMG%":0.0,
}
HP_SUP_W = {   # Mika, Selena, Tressa
    "HP":1.0,"HP%":0.9,"DEF":0.8,"DEF%":0.7,"EGO Recovery":0.5,
    "Crit Rate%":0.0,"Crit DMG%":0.0,"ATK":0.0,"ATK%":0.0,
    "DoT%":0.0,"Extra DMG%":0.0,
}

# ── Characters ────────────────────────────────────────────────────────────────
CHARS = {
    "Nine":     {"sets":["Line of Justice","Beast's Yearning","Conqueror's Aspect",
                         "Executioner's Tool","Blackwing","Cursed Corpse","Bullet of Order"],
                 "weights":NINE_W,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%",
                              "ideal":"Crit Rate%"},
                 "goal":"350 DEF · Potential 7"},
    "Veronica": {"sets":["Line of Justice","Beast's Yearning","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%",
                              "ideal":"Crit Rate%"},
                 "goal":"Crit DPS"},
    "Beryl":    {"sets":["Line of Justice","Executioner's Tool","Blackwing","Cursed Corpse"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%",
                              "ideal":"Crit Rate%"},
                 "goal":"Burst DPS"},
    "Sereniel": {"sets":["Executioner's Tool","Blackwing","Instinctual Growth",
                         "Judgment's Flames","Cursed Corpse"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Instinct DMG%","imagination":"Crit Rate%"},
                 "goal":"60% Crit Rate · Potential 7"},
    "Rin":      {"sets":["Conqueror's Aspect","Orb of Inhibition","Executioner's Tool","Blackwing",
                         "Offering of the Void"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Void DMG%","imagination":"Crit Rate%"},
                 "goal":"Void / 1-AP DPS"},
    "Renoa":    {"sets":["Conqueror's Aspect","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%"},
                 "goal":"1-AP card DPS"},
    "Chizuru":  {"sets":["Orb of Inhibition","Executioner's Tool","Seth's Scarab","Blackwing",
                         "Offering of the Void"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Void DMG%","imagination":"Crit Rate%"},
                 "goal":"Void Multi-Hit DPS"},
    "Tiphera":  {"sets":["Line of Justice","Tetra's Authority","Executioner's Tool","Healer's Journey"],
                 "weights":TIPHERA_W,
                 "good_main":{"desire":"Order DMG%","imagination":"DEF%"},
                 "goal":"DEF scaler · 350 DEF"},
    "Hugo":     {"sets":["Executioner's Tool","Blackwing","Line of Justice"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%"},
                 "goal":"Follow-up DPS"},
    "Khalipe":  {"sets":["Instinctual Growth","Tetra's Authority","Executioner's Tool"],
                 "weights":KHALIPE_W,
                 "good_main":{"desire":"Instinct DMG%","imagination":"DEF%"},
                 "goal":"300 DEF + Crit · Potential 7"},
    "Orlea":    {"sets":["Glory's Reign","Tetra's Authority","Healer's Journey","Instinctual Growth"],
                 "weights":DEF_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"300 DEF · Potential 7"},
    "Cassius":  {"sets":["Glory's Reign","Tetra's Authority","Healer's Journey","Seth's Scarab"],
                 "weights":DEF_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"DEF + HP support"},
    "Mika":     {"sets":["Tetra's Authority","Healer's Journey","Seth's Scarab"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"Healer / Shield support"},
    "Selena":   {"sets":["Healer's Journey","Tetra's Authority","Seth's Scarab"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"Sub-DPS support"},
    "Tressa":   {"sets":["Cursed Corpse","Healer's Journey","Tetra's Authority","Seth's Scarab"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"Agony support"},
    # ── New characters ────────────────────────────────────────────────────────
    # Amir — DEF-scaler (Metalization), 301 DEF goal
    "Amir":     {"sets":["Conqueror's Aspect","Tetra's Authority","Executioner's Tool",
                         "Blackwing","Cursed Corpse"],
                 "weights":{"Crit Rate%":0.75,"Crit DMG%":0.75,"DEF":1.0,"DEF%":0.8,
                             "ATK":0.1,"ATK%":0.1,"HP":0.0,"HP%":0.0,"EGO Recovery":0.3,
                             "DoT%":0.0,"Extra DMG%":0.0},
                 "good_main":{"desire":"Void DMG%","imagination":"DEF%"},
                 "goal":"301 DEF · Metalization build"},
    # Diana — new Void DPS
    "Diana":    {"sets":["Orb of Inhibition","Conqueror's Aspect","Executioner's Tool",
                         "Blackwing","Offering of the Void"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Void DMG%","imagination":"Crit Rate%"},
                 "goal":"Void Discard DPS"},
    # Haru — Justice Crit DPS
    "Haru":     {"sets":["Line of Justice","Conqueror's Aspect","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Justice DMG%","imagination":"Crit Rate%"},
                 "goal":"Justice Crit DPS"},
    # Kayron — Void DoT DPS
    "Kayron":   {"sets":["Offering of the Void","Orb of Inhibition","Cursed Corpse",
                         "Executioner's Tool","Blackwing"],
                 "weights":{"Crit Rate%":0.75,"Crit DMG%":0.75,"DoT%":1.0,"Extra DMG%":0.5,
                             "ATK":0.3,"ATK%":0.3,"DEF":0.0,"DEF%":0.0,"HP":0.0,
                             "HP%":0.0,"EGO Recovery":0.0},
                 "good_main":{"desire":"Void DMG%","imagination":"Crit Rate%"},
                 "goal":"Void DoT DPS"},
    # Lucas — Agony/DoT sub-DPS
    "Lucas":    {"sets":["Cursed Corpse","Conqueror's Aspect","Executioner's Tool","Seth's Scarab"],
                 "weights":{"Crit Rate%":0.75,"Crit DMG%":0.75,"DoT%":0.75,"ATK":0.4,
                             "ATK%":0.4,"DEF":0.0,"DEF%":0.0,"HP":0.0,"HP%":0.0,
                             "EGO Recovery":0.0,"Extra DMG%":0.0},
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%"},
                 "goal":"Agony sub-DPS"},
    # Luke — Order Bullet DPS
    "Luke":     {"sets":["Bullet of Order","Line of Justice","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%"},
                 "goal":"Order Bullet DPS"},
    # Magna — Shield/Counter Support-DPS
    "Magna":    {"sets":["Tetra's Authority","Conqueror's Aspect","Executioner's Tool","Blackwing"],
                 "weights":{"DEF":1.0,"DEF%":0.8,"Crit Rate%":0.6,"Crit DMG%":0.6,
                             "HP":0.2,"HP%":0.2,"ATK":0.0,"ATK%":0.0,"EGO Recovery":0.2,
                             "DoT%":0.0,"Extra DMG%":0.0},
                 "good_main":{"desire":"Void DMG%","imagination":"DEF%"},
                 "goal":"DEF + Crit Counter build"},
    # Maribell — Shield/DPS (scales off team shields)
    "Maribell": {"sets":["Tetra's Authority","Spark of Passion","Executioner's Tool","Blackwing"],
                 "weights":{"DEF":0.8,"DEF%":0.7,"Crit Rate%":0.7,"Crit DMG%":0.7,
                             "HP":0.1,"HP%":0.1,"ATK":0.1,"ATK%":0.1,"EGO Recovery":0.2,
                             "DoT%":0.0,"Extra DMG%":0.0},
                 "good_main":{"desire":"Passion DMG%","imagination":"DEF%"},
                 "goal":"Shield-scaling DPS"},
    # Mei Lin — Passion DPS
    "Mei Lin":  {"sets":["Spark of Passion","Executioner's Tool","Blackwing","Conqueror's Aspect"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Passion DMG%","imagination":"Crit Rate%"},
                 "goal":"Passion Upgrade DPS"},
    # Narja — AP Support
    "Narja":    {"sets":["Tetra's Authority","Healer's Journey","Seth's Scarab"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"AP cost reduction support"},
    # Nia — Draw/Discard Support
    "Nia":      {"sets":["Healer's Journey","Tetra's Authority","Seth's Scarab"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"Draw/Discard engine support"},
    # Owen — Heal Support
    "Owen":     {"sets":["Healer's Journey","Tetra's Authority","Seth's Scarab","Glory's Reign"],
                 "weights":HP_SUP_W,
                 "good_main":{"desire":"HP%","imagination":"DEF%"},
                 "goal":"Heal support"},
    # Rei — Morale/1-cost Support-DPS
    "Rei":      {"sets":["Offering of the Void","Cursed Corpse","Executioner's Tool",
                         "Seth's Scarab","Healer's Journey"],
                 "weights":{"Crit Rate%":0.6,"Crit DMG%":0.6,"DEF":0.5,"DEF%":0.5,
                             "HP":0.5,"HP%":0.5,"ATK":0.2,"ATK%":0.2,"EGO Recovery":0.4,
                             "DoT%":0.0,"Extra DMG%":0.0},
                 "good_main":{"desire":"Void DMG%","imagination":"DEF%"},
                 "goal":"Morale support / sub-DPS"},
    # Rita — Instinct DPS (new)
    "Rita":     {"sets":["Instinctual Growth","Judgment's Flames","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Instinct DMG%","imagination":"Crit Rate%"},
                 "goal":"Instinct DPS"},
    # Yuki — Order DPS
    "Yuki":     {"sets":["Bullet of Order","Line of Justice","Executioner's Tool","Blackwing"],
                 "weights":CRIT_DPS,
                 "good_main":{"desire":"Order DMG%","imagination":"Crit Rate%"},
                 "goal":"Order DPS"},
}

SET_TO_CHARS = {}  # dict[str, list[str]]
for _c, _d in CHARS.items():
    for _s in _d["sets"]:
        SET_TO_CHARS.setdefault(_s,[]).append(_c)

SET_ALIASES = {
    "spark":"Spark of Passion","spark of passion":"Spark of Passion",
    "offering of the void":"Offering of the Void","offering":"Offering of the Void","void offering":"Offering of the Void",
    "bullet of order":"Bullet of Order","bullet":"Bullet of Order",
    "glory's reign":"Glory's Reign","glory's":"Glory's Reign","glory":"Glory's Reign",
    "executioner's tool":"Executioner's Tool","executioner's":"Executioner's Tool","executioner":"Executioner's Tool",
    "blackwing":"Blackwing","black wing":"Blackwing",
    "tetra's authority":"Tetra's Authority","tetra's":"Tetra's Authority","tetra":"Tetra's Authority",
    "healer's journey":"Healer's Journey","healer's":"Healer's Journey","healer":"Healer's Journey",
    "seth's scarab":"Seth's Scarab","seth's":"Seth's Scarab","seth":"Seth's Scarab","scarab":"Seth's Scarab",
    "cursed corpse":"Cursed Corpse","cursed":"Cursed Corpse","corpse":"Cursed Corpse",
    "conqueror's aspect":"Conqueror's Aspect","conqueror's":"Conqueror's Aspect","conqueror":"Conqueror's Aspect",
    "judgment's flames":"Judgment's Flames","judgment's":"Judgment's Flames","judgment":"Judgment's Flames","flames":"Judgment's Flames",
    "orb of inhibition":"Orb of Inhibition","orb":"Orb of Inhibition","inhibition":"Orb of Inhibition",
    "instinctual growth":"Instinctual Growth","instinctual":"Instinctual Growth","instinct":"Instinctual Growth",
    "line of justice":"Line of Justice","line":"Line of Justice","justice":"Line of Justice",
    "beast's yearning":"Beast's Yearning","beast's":"Beast's Yearning","beast":"Beast's Yearning","yearning":"Beast's Yearning",
}

def find_set(text: str):
    t = text.lower()
    for alias in sorted(SET_ALIASES, key=len, reverse=True):
        if alias in t:
            return SET_ALIASES[alias]
    return None

def find_slot(text: str):
    """Detect slot from a line that looks like a slot label.
    Rejects lines with set keywords or longer than 30 chars.
    """
    t = text.strip().lower()
    set_words = ["of the", "yearning", "justice", "executioner", "conqueror",
                 "offering", "instinct", "judgment", "scarab", "corpse",
                 "blackwing", "glory", "tetra", "healer", "orb", "bullet",
                 "spark", "passion", "beast", "line", "void offering"]
    for sw in set_words:
        if sw in t:
            return None
    if len(text.strip()) > 30:
        return None
    for alias, slot in SLOT_NAMES.items():
        if re.search(r'\b' + re.escape(alias) + r'\b', t):
            return slot
    return None

def find_slot_from_title(text: str):
    """Extract slot from the fragment title line.
    Fragment titles end with the slot type: '...Shock', '...Suppression', '...Denial',
    '...Ideal', '...Desire', '...Imagination', or a variant name like 'Longing'/'Anomaly'.
    Returns the slot key if found.
    """
    t = text.strip().lower()
    # Only match if this looks like a fragment title (contains a set name)
    is_title = any(sw in t for sw in ["of the", "yearning", "justice", "executioner",
                                       "conqueror", "offering", "instinct", "judgment",
                                       "scarab", "corpse", "blackwing", "glory", "tetra",
                                       "healer", "orb", "bullet", "spark", "beast", "wing"])
    if not is_title:
        return None
    # The slot type is the last word (or last two words) of the title
    # Map known slot suffixes
    TITLE_SLOT_MAP = {
        "shock":        "shock",
        "suppression":  "suppression",
        "denial":       "denial",
        "ideal":        "ideal",
        "desire":       "desire",
        "imagination":  "imagination",
        # Variant names used in-game for the same slot types
        "longing":      "ideal",    # e.g. "Black Wing Longing" = Ideal slot
        "anomaly":      "ideal",
        "resolution":   "desire",
        "fragment":     "imagination",
        "trace":        "imagination",
    }
    words = t.split()
    # Check last 1-2 words
    for n in (1, 2):
        suffix = " ".join(words[-n:]) if len(words) >= n else ""
        if suffix in TITLE_SLOT_MAP:
            return TITLE_SLOT_MAP[suffix]
    return None

# ── Stat normalisation ────────────────────────────────────────────────────────
STAT_MAP = [
    (r"critical\s*(chance|rate|hit)",  "Crit Rate%"),
    (r"crit\s*(chance|rate|hit)",       "Crit Rate%"),
    (r"critical\s*damage",              "Crit DMG%"),
    (r"crit\s*(damage|dmg)",            "Crit DMG%"),
    (r"ego\s*recov\w*",                 "EGO Recovery"),
    (r"damage\s*over\s*time",           "DoT%"),
    (r"extra\s*damage",                 "Extra DMG%"),
    (r"justice\s*damage",               "Justice DMG%"),
    (r"order\s*damage",                 "Order DMG%"),
    (r"void\s*damage",                  "Void DMG%"),
    (r"instinct\s*damage",              "Instinct DMG%"),
    (r"chaos\s*damage",                 "Chaos DMG%"),
    (r"passion\s*damage",               "Passion DMG%"),
    (r"defense|defence",                "DEF"),
    (r"\battack\b",                     "ATK"),
    # Health: broad pattern to catch OCR misreads like "Heatth", "Heaith", "Hea|th"
    (r"h[e3][a@][l1|i][t7][h#]|health|\bhp\b", "HP"),
]

# Attribute damage types — always Main-Stats on variable slots (IV, V, VI)
ATTR_DMG_STATS = {
    "Order DMG%","Void DMG%","Instinct DMG%","Chaos DMG%",
    "Justice DMG%","Passion DMG%","Attribute DMG%"
}

def normalize_stat(raw: str) -> str:
    r = raw.lower().strip()
    has_pct = "%" in r
    r = r.replace("%","").strip()
    for pat, name in STAT_MAP:
        if re.search(pat, r):
            if has_pct and not name.endswith("%"):
                return name + "%"
            return name
    return raw.strip().title()

def parse_value(s: str) -> float:
    try:
        s = s.replace(",", ".")
        return float(re.sub(r"[^0-9.\-]", "", s))
    except:
        return 0.0

# ── OCR ───────────────────────────────────────────────────────────────────────
STAT_LINE = re.compile(
    r'(justice\s*damage|order\s*damage|void\s*damage|instinct\s*damage|'
    r'chaos\s*damage|passion\s*damage|'
    r'critical\s*(?:chance|rate|hit|damage)|crit\s*(?:chance|rate|hit|damage|dmg)|'
    r'ego\s*recov\w*|damage\s*over\s*time|extra\s*damage|defense|defence|attack|'
    # Health: catch common OCR misreads (l→1, l→|, l→i, e→3, a→@)
    r'h[e3][a@][l1|i][t7][h#]|health)'
    r'(.*)',
    re.IGNORECASE
)

def extract_best_value(rest: str) -> str:
    """Extract best value from rest of stat line.
    Handles: '+3 3.9%', '+1 +7.2%', '+7.2%', '+14', '+3,0%' (comma decimal)
    Prefers percentage over flat value.
    """
    # Normalise comma decimals first
    rest = rest.replace(",", ".")
    tokens = re.findall(r'[+\-]?\d+\.?\d*\s*%?', rest)
    if not tokens:
        return "0"
    pct_vals = [t for t in tokens if "%" in t]
    if pct_vals:
        return pct_vals[-1].strip()
    return tokens[-1].strip()

def fix_ocr_text(text: str) -> str:
    """Fix common Tesseract misreads before parsing."""
    # Health misreads — Tesseract often reads l as 1, |, or i
    text = re.sub(r'\bHea[l1|i][t7][h#]\b', 'Health', text, flags=re.IGNORECASE)
    text = re.sub(r'\bHea[l1|i]th\b',       'Health', text, flags=re.IGNORECASE)
    text = re.sub(r'\bHeatth\b',             'Health', text, flags=re.IGNORECASE)
    text = re.sub(r'\bHeaith\b',             'Health', text, flags=re.IGNORECASE)
    text = re.sub(r'\bHea1th\b',             'Health', text, flags=re.IGNORECASE)
    # Defense misreads
    text = re.sub(r'\bDefanse\b',            'Defense', text, flags=re.IGNORECASE)
    text = re.sub(r'\bDefence\b',            'Defense', text, flags=re.IGNORECASE)
    # Attack misreads
    text = re.sub(r'\bAttaok\b',             'Attack', text, flags=re.IGNORECASE)
    text = re.sub(r'\bAttask\b',             'Attack', text, flags=re.IGNORECASE)
    # Critical misreads
    text = re.sub(r'\bCritica[l1]\b',        'Critical', text, flags=re.IGNORECASE)
    text = re.sub(r'\bCrit[i1]cal\b',        'Critical', text, flags=re.IGNORECASE)
    # Damage misreads
    text = re.sub(r'\bDamaqe\b',             'Damage', text, flags=re.IGNORECASE)
    text = re.sub(r'\bDarnage\b',            'Damage', text, flags=re.IGNORECASE)
    return text

def parse_fragment(text: str) -> dict:
    text  = fix_ocr_text(text)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    r = {"set":None,"slot":None,"rarity":None,"upgrade":None,
         "main_stat":None,"sub_stats":[],"raw_lines":lines}

    for line in lines:
        if not r["set"]:
            r["set"] = find_set(line)
            if r["set"] and not r["slot"]:
                r["slot"] = find_slot_from_title(line)
        elif not r["slot"]:
            # Try title-based detection on every line even after set is found
            if not find_set(line) is None or True:
                s = find_slot_from_title(line)
                if s: r["slot"] = s
        if not r["slot"]:  r["slot"] = find_slot(line)
        if not r["rarity"]:
            for x in ["legendary","epic","rare","common"]:
                if x in line.lower(): r["rarity"]=x.title(); break
        if not r["upgrade"]:
            m = re.search(r'\+(\d)\b',line)
            if m: r["upgrade"]="+"+m.group(1)

    # Override: if title-based slot was found, always use it
    # It's more reliable than the OCR label (which can pick up the overlay window)
    for line in lines:
        s = find_slot_from_title(line)
        if s:
            r["slot"] = s
            break

    seen, entries = set(), []
    for line in lines:
        m = STAT_LINE.search(line)
        if m:
            raw_name = m.group(1)
            rest     = m.group(2)
            val      = extract_best_value(rest)
            if not val or val in ("0",""):
                continue

            # Determine if this is flat or % variant
            # e.g. "Health 1.3%" → HP%,  "Health +11" → HP
            # "Defense +3 3.9%" → DEF%,  "Defense +5" → DEF
            is_pct = "%" in val
            base_name = normalize_stat(raw_name)

            # For stats that have both flat and % variants, distinguish them
            if base_name in ("HP","DEF","ATK") and is_pct:
                name = base_name + "%"
            else:
                name = base_name

            # Use name as dedup key — allows both HP and HP% on same fragment
            if name not in seen:
                seen.add(name)
                entries.append({"name":name,"value":val,"fval":parse_value(val)})

    if entries:
        slot  = r.get("slot")
        fixed = FIXED_MAIN.get(slot) if slot else None

        # Attribute DMG% stats (Justice/Order/Void/etc.) are ALWAYS the main stat
        # — they never appear as sub-stats, so if we see one, it must be the main
        attr_main = None
        for e in entries:
            if e["name"] in ATTR_DMG_STATS:
                attr_main = e
                break

        if attr_main:
            # Attribute DMG% (Justice/Order/Void etc.) — always main, no ambiguity
            r["main_stat"] = attr_main
            r["sub_stats"] = [e for e in entries if e is not attr_main]
        elif fixed:
            # Fixed slot (Shock=ATK, Suppression=DEF, Denial=HP) — find by name
            match = next((e for e in entries if e["name"] == fixed), None)
            if match:
                r["main_stat"] = match
                r["sub_stats"] = [e for e in entries if e is not match]
            else:
                # Fixed stat not found — trust first entry
                r["main_stat"] = entries[0]
                r["sub_stats"] = entries[1:]
        else:
            # Variable slot (Ideal/Desire/Imagination) or slot unknown:
            # First recognised stat in the OCR text = Main-Stat.
            # This works because Tesseract reads top-to-bottom,
            # and the game always shows the Main-Stat at the top (orange).
            r["main_stat"] = entries[0]
            r["sub_stats"] = entries[1:]
    return r

# ── Per-character scoring ────────────────────────────────────────────────────
def score_for_char(char_name, subs, main_name, slot):
    """Score a fragment for one specific character."""
    cd = CHARS[char_name]
    cw = cd["weights"]
    top4_w = sorted(cw.values(), reverse=True)[:4]
    ideal  = sum(x * 100.0 for x in top4_w) or 400.0

    sub_score = 0.0
    details   = []
    for s in subs:
        sw  = cw.get(s["name"], 0.0)
        rv  = roll_value(s["name"], s["fval"])
        contrib = sw * (70.0 + rv * 30.0)
        sub_score += contrib
        details.append({"name":s["name"],"value":s["value"],"w":sw,"rv":rv,"contrib":contrib})

    raw_pct = (sub_score / ideal) * 100.0

    main_ok    = None
    main_bonus = 0.0
    if slot and main_name:
        fixed = FIXED_MAIN.get(slot)
        if fixed:
            main_ok = (main_name == fixed)
            if not main_ok: main_bonus = -10.0
        else:
            good_main = cd.get("good_main", {}).get(slot)
            if good_main:
                main_ok = (main_name == good_main)
                main_bonus = +5.0 if main_ok else -8.0

    score = round(max(0.0, min(100.0, raw_pct + main_bonus)))
    return {"char":char_name,"score":score,"details":details,
            "main_ok":main_ok,"main_bonus":main_bonus}


def score_fragment(frag):
    set_key   = frag.get("set")
    slot      = frag.get("slot")
    subs      = frag.get("sub_stats", [])
    main      = frag.get("main_stat")
    main_name = main["name"] if main else None

    if not set_key:
        return {"score":0,"grade":"?","verdict":"Set not recognised","color":C["muted"],
                "details":[],"main_ok":None,"chars":[],"set_found":False,
                "best_char":None,"char_scores":[]}

    chars = SET_TO_CHARS.get(set_key, [])
    if not chars:
        # Set name recognised but no char data — show stats without scoring
        details = [{"name":s["name"],"value":s["value"],"w":0.0,
                    "rv":roll_value(s["name"],s["fval"]),"contrib":0} for s in subs]
        return {"score":0,"grade":"?","verdict":"No data for this set","color":C["muted"],
                "details":details,"main_ok":None,"chars":[],"set_found":True,
                "best_char":None,"char_scores":[]}

    # Score independently per character, take the best match
    char_scores = sorted(
        [score_for_char(ch, subs, main_name, slot) for ch in chars],
        key=lambda x: -x["score"]
    )
    best = char_scores[0]

    score   = best["score"]
    details = best["details"]
    main_ok = best["main_ok"]

    if score >= 75:   grade,color,verdict = "S", C["s"], "Keep"
    elif score >= 55: grade,color,verdict = "A", C["a"], "Keep"
    elif score >= 38: grade,color,verdict = "B", C["b"], "Maybe"
    elif score >= 22: grade,color,verdict = "C", C["c"], "Discard"
    elif score >= 10: grade,color,verdict = "D", C["d"], "Discard"
    else:             grade,color,verdict = "F", C["f"], "Discard"

    return {"score":score,"grade":grade,"verdict":verdict,"color":color,
            "details":details,"main_ok":main_ok,"chars":chars,"set_found":True,
            "best_char":best["char"],"char_scores":char_scores,
            "main_bonus":best.get("main_bonus",0)}


# ── Overlay ───────────────────────────────────────────────────────────────────
class CZNOverlay:
    SZ = 56
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CZN Overlay")
        # Don't use overrideredirect on first show — causes invisible window on some systems
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.92)
        self.root.configure(bg=C["bg"])
        self.root.resizable(False, False)
        # Remove title bar but keep window visible
        self.root.overrideredirect(True)
        self._dx = self._dy = 0
        self._rwin = None
        self._oval_id = None
        self._build_btn()
        sw, sh = pyautogui.size()
        self.root.geometry(f"{self.SZ}x{self.SZ}+{sw - self.SZ - 20}+{sh // 2}")
        self.root.deiconify()
        self.root.lift()
        self.root.update()
        _hook.add(VK_F9,     self._scan)
        _hook.add(VK_ESCAPE, self._quit)
        _hook.start()

    def _build_btn(self):
        cv = tk.Canvas(self.root, width=self.SZ, height=self.SZ,
                       bg=C["bg"], highlightthickness=0)
        cv.pack(fill="both", expand=True)
        p = 3
        self._oval_id = cv.create_oval(
            p, p, self.SZ-p, self.SZ-p,
            fill=C["p1_bg"], outline=C["p1_fg"], width=2
        )
        cv.create_text(self.SZ//2, self.SZ//2 - 5,
                       text="CZN", fill=C["p1_fg"],
                       font=("Segoe UI", 10, "bold"))
        cv.create_text(self.SZ//2, self.SZ - 10,
                       text="F9", fill=C["p1_fg"],
                       font=("Segoe UI", 8))
        cv.bind("<Button-1>",      self._click)
        cv.bind("<ButtonPress-1>", lambda e: (setattr(self, '_dx', e.x), setattr(self, '_dy', e.y)))
        cv.bind("<B1-Motion>",     self._drag)
        cv.bind("<Button-3>",      lambda e: self._quit())
        self._cv = cv

    def _drag(self, e):
        x = self.root.winfo_x() + e.x - self._dx
        y = self.root.winfo_y() + e.y - self._dy
        self.root.geometry(f"+{x}+{y}")

    def _click(self, e=None):
        threading.Thread(target=self._scan, daemon=True).start()

    def _scan(self):
        if self._oval_id:
            self.root.after(0, lambda: self._cv.itemconfig(self._oval_id, fill="#1a4080"))
        try:
            sw, sh = pyautogui.size()

            # Scan right 55% of screen, skip bottom 15% (fragment thumbnails row)
            # The detail panel is always on the right side of the screen
            # Skipping the bottom avoids picking up thumbnail text and the overlay itself
            x1 = int(sw * 0.38)
            y1 = int(sh * 0.03)
            x2 = int(sw * 0.98)
            y2 = int(sh * 0.83)  # cut off thumbnail strip at bottom

            img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            base = os.path.dirname(__file__)
            img.save(os.path.join(base, "debug_scan.png"))
            ocr = run_ocr(img)
            with open(os.path.join(base, "debug_ocr.txt"), "w", encoding="utf-8") as fh:
                fh.write(ocr)
            frag = parse_fragment(ocr); frag["ocr_raw"] = ocr
            res  = score_fragment(frag)
            self.root.after(0, lambda: self._show(frag, res))
        except Exception as ex:
            self.root.after(0, lambda: self._err(str(ex)))
        finally:
            if self._oval_id:
                self.root.after(0, lambda: self._cv.itemconfig(self._oval_id, fill=C["p1_bg"]))

    def _show(self,frag,res):
        if self._rwin and self._rwin.winfo_exists():
            self._rwin.destroy()
        win=tk.Toplevel(self.root)
        win.overrideredirect(True); win.attributes("-topmost",True)
        win.attributes("-alpha",0.96); win.configure(bg=C["bg"])
        f=tk.Frame(win,bg=C["bg"],padx=14,pady=12); f.pack(fill="both",expand=True)

        # ── Header ──
        hdr=tk.Frame(f,bg=C["bg"]); hdr.pack(fill="x",pady=(0,6))
        tk.Label(hdr,text=frag.get("set") or "Unknown Set",bg=C["bg"],fg=C["text"],
                 font=("Segoe UI",12,"bold")).pack(side="left")
        slot=frag.get("slot")
        if slot:
            tk.Label(hdr,text=SLOT_DISPLAY.get(slot,slot),bg=C["bg"],fg=C["muted"],
                     font=("Segoe UI",9)).pack(side="left",padx=8)
        info=" ".join(filter(None,[frag.get("rarity"),frag.get("upgrade")]))
        if info: tk.Label(hdr,text=info,bg=C["bg"],fg=C["muted"],font=("Segoe UI",9)).pack(side="right")

        # ── Score row ──
        sr=tk.Frame(f,bg=C["bg"]); sr.pack(fill="x",pady=(0,6))

        # Big score
        tk.Label(sr,text=f"{res['score']}%",bg=C["bg"],fg=res["color"],
                 font=("Segoe UI",30,"bold")).pack(side="left")

        # Grade badge
        gbg = res["color"]; gfg = C["bg"] if res["grade"] not in ("F",) else C["muted"]
        tk.Label(sr,text=res["grade"],bg=gbg,fg=C["bg"],
                 font=("Segoe UI",16,"bold"),width=2,pady=2).pack(side="left",padx=6)

        # Verdict badge
        vbg=C["p1_bg"] if res["verdict"]=="Keep" else \
            C["p2_bg"] if res["verdict"]=="Maybe" else C["bad_bg"]
        vfg=C["p1_fg"] if res["verdict"]=="Keep" else \
            C["p2_fg"] if res["verdict"]=="Maybe" else "#888"
        tk.Label(sr,text=res["verdict"],bg=vbg,fg=vfg,
                 font=("Segoe UI",11,"bold"),padx=10,pady=4).pack(side="left",padx=4)

        tk.Frame(f,bg=C["border"],height=1).pack(fill="x",pady=(2,6))

        # ── Main stat ──
        main=frag.get("main_stat")
        if main:
            mr=tk.Frame(f,bg=C["bg"]); mr.pack(fill="x",pady=2)
            mok=res.get("main_ok")
            if mok is True:    mbg,mfg,mlbl=C["main_ok_bg"],C["main_ok_fg"],"Main ✓"
            elif mok is False: mbg,mfg,mlbl=C["main_bad_bg"],C["main_bad_fg"],"Main ✗"
            else:              mbg,mfg,mlbl=C["main_neu_bg"],C["main_neu_fg"],"Main"
            tk.Label(mr,text=mlbl,bg=mbg,fg=mfg,font=("Segoe UI",8,"bold"),
                     padx=4,pady=1).pack(side="left")
            tk.Label(mr,text=f"  {main['name']}",bg=C["bg"],fg=C["text"],
                     font=("Segoe UI",10)).pack(side="left")
            tk.Label(mr,text=main["value"],bg=C["bg"],fg=C["muted"],
                     font=("Segoe UI",10)).pack(side="right")
            tk.Frame(f,bg=C["border"],height=1).pack(fill="x",pady=(4,4))

        # ── Sub stats ──
        details=res.get("details",[])
        if details:
            for d in details:
                row=tk.Frame(f,bg=C["bg"]); row.pack(fill="x",pady=2)
                w=d["w"]; rv=d["rv"]
                # Weight badge colour
                if w>=0.9:   lbg,lfg=C["p1_bg"],C["p1_fg"]
                elif w>=0.6: lbg,lfg=C["p2_bg"],C["p2_fg"]
                elif w>=0.3: lbg,lfg="#1a2a1a","#9ecf9e"
                else:        lbg,lfg=C["bad_bg"],C["bad_fg"]
                # Roll value indicator (e.g. "87%")
                rv_str = f"{round(rv*100)}%" if w>0 else "–"
                tk.Label(row,text=rv_str,bg=lbg,fg=lfg,font=("Segoe UI",9,"bold"),
                         width=5,padx=3,pady=1).pack(side="left")
                tk.Label(row,text=f"  {d['name']}",bg=C["bg"],fg=C["text"],
                         font=("Segoe UI",10)).pack(side="left")
                tk.Label(row,text=d["value"],bg=C["bg"],fg=C["muted"],
                         font=("Segoe UI",10)).pack(side="right")
        else:
            # Debug fallback
            raw=frag.get("ocr_raw","")
            tk.Label(f,text="Stats not recognised — OCR text:",bg=C["bg"],fg="#ff9500",
                     font=("Segoe UI",9,"bold")).pack(anchor="w")
            tk.Label(f,text="\n".join(raw.strip().splitlines()[:10]) or "(empty)",
                     bg=C["bg"],fg=C["muted"],font=("Courier New",8),
                     justify="left",wraplength=260).pack(anchor="w")
            tk.Label(f,text="→ debug_scan.png + debug_ocr.txt saved",
                     bg=C["bg"],fg=C["muted"],font=("Segoe UI",7)).pack(anchor="w",pady=(2,0))

        # ── Per-character scores ──
        char_scores = res.get("char_scores", [])
        best_char   = res.get("best_char")
        if char_scores:
            tk.Frame(f, bg=C["border"], height=1).pack(fill="x", pady=(6,4))
            # Best char highlighted, rest smaller
            for cs in char_scores:
                cr = tk.Frame(f, bg=C["bg"]); cr.pack(fill="x", pady=1)
                is_best = cs["char"] == best_char
                # Score badge
                if cs["score"] >= 55:   sbg,sfg = C["p1_bg"],C["p1_fg"]
                elif cs["score"] >= 38: sbg,sfg = C["p2_bg"],C["p2_fg"]
                else:                   sbg,sfg = C["bad_bg"],C["bad_fg"]
                tk.Label(cr, text=f"{cs['score']}%", bg=sbg, fg=sfg,
                         font=("Segoe UI", 8 if is_best else 7, "bold"),
                         width=5, padx=2).pack(side="left")
                tk.Label(cr, text=f"  {cs['char']}",
                         bg=C["bg"],
                         fg=C["text"] if is_best else C["muted"],
                         font=("Segoe UI", 10 if is_best else 9,
                               "bold" if is_best else "normal")).pack(side="left")
                if cs.get("main_ok") is False:
                    tk.Label(cr, text="main ✗", bg=C["bg"], fg=C["main_bad_fg"],
                             font=("Segoe UI",8)).pack(side="right")

        # ── Close ──
        tk.Frame(f,bg=C["border"],height=1).pack(fill="x",pady=(8,3))
        tk.Button(f,text="✕  Close",command=win.destroy,bg="#1f2937",fg=C["muted"],
                  relief="flat",font=("Segoe UI",9),cursor="hand2",bd=0).pack(side="right")

        win.update_idletasks()
        bx,by=self.root.winfo_x(),self.root.winfo_y()
        sw,_=pyautogui.size()
        wx=bx-win.winfo_width()-8 if bx>sw//2 else bx+self.SZ+8
        win.geometry(f"+{wx}+{max(10,by-80)}")
        self._rwin=win

    def _err(self,msg):
        if self._rwin and self._rwin.winfo_exists(): self._rwin.destroy()
        win=tk.Toplevel(self.root); win.overrideredirect(True)
        win.attributes("-topmost",True); win.configure(bg=C["bg"])
        f=tk.Frame(win,bg=C["bg"],padx=14,pady=12); f.pack()
        tk.Label(f,text=f"Error:\n{msg[:120]}",bg=C["bg"],fg="#ff6b6b",
                 font=("Segoe UI",9),wraplength=260).pack()
        tk.Button(f,text="OK",command=win.destroy,bg="#1f2937",fg=C["text"],relief="flat").pack(pady=6)
        bx,by=self.root.winfo_x(),self.root.winfo_y()
        win.update_idletasks()
        win.geometry(f"+{max(0,bx-280)}+{by}")
        self._rwin=win

    def _quit(self):
        _hook.stop()
        self.root.destroy()

    def run(self):
        print("CZN Overlay v4 — F9=scan | Right-click=quit")
        self.root.mainloop()

if __name__ == "__main__":
    CZNOverlay().run()
