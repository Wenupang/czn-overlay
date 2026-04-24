"""
czn_overlay.py — CZN Fragment Rater v5
========================================
Hybrid overlay: F9 or "Scan" button triggers OCR to pre-fill fields.
All fields are editable. Upgrade level always manual.
Drag title bar to reposition. Click X to hide (keeps running).
"""

import tkinter as tk
from tkinter import ttk
import threading, os, sys, json, re

try:
    import pyautogui
    from PIL import ImageGrab, Image
except ImportError:
    print("Missing: pip install pyautogui Pillow"); sys.exit(1)

from data    import CHARS, SET_TO_CHARS, FIXED_MAIN, COLORS as C, ROLL_RANGES, DREAM_MAX, SET_ALIASES
from scoring import score_fragment

# ── Constants ─────────────────────────────────────────────────────────────────
SETS = sorted({s for char in CHARS.values() for s in char["sets"]})

ALL_SUBSTATS = [
    "", "ATK", "DEF", "HP",
    "ATK%", "DEF%", "HP%",
    "Crit Rate%", "Crit DMG%",
    "EGO Recovery", "DoT%", "Extra DMG%",
]

SLOTS = [
    ("shock",       "I · Shock"),
    ("suppression", "II · Suppression"),
    ("denial",      "III · Denial"),
    ("ideal",       "IV · Ideal"),
    ("desire",      "V · Desire"),
    ("imagination", "VI · Imagination"),
]
SLOT_LABEL_TO_KEY = {v: k for k, v in SLOTS}

MAIN_BY_SLOT = {
    "shock":       ["ATK"],
    "suppression": ["DEF"],
    "denial":      ["HP"],
    "ideal":       ["Crit Rate%","Crit DMG%","ATK%","HP%","DEF%"],
    "desire":      ["Justice DMG%","Order DMG%","Void DMG%",
                    "Instinct DMG%","Passion DMG%","Chaos DMG%","ATK%","HP%"],
    "imagination": ["Crit Rate%","Crit DMG%","ATK%","HP%","DEF%","EGO Recovery"],
}

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")

# ── OCR helpers ───────────────────────────────────────────────────────────────
def _run_ocr(img: Image.Image) -> str:
    """Windows OCR via PowerShell, Tesseract fallback."""
    import tempfile, subprocess
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        path = f.name
    img.save(path, "PNG")
    ps = f"""
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null=[Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
function Aw($t,$r){{([System.WindowsRuntimeSystemExtensions].GetMethods()|?{{$_.Name-eq'AsTask'-and$_.GetParameters().Count-eq 1-and!$_.IsGenericMethod}}|select -f 1).MakeGenericMethod($r).Invoke($null,@($t)).GetAwaiter().GetResult()}}
$file=Aw ([Windows.Storage.StorageFile]::GetFileFromPathAsync('{path.replace(chr(92),chr(92)*2)}')) ([Windows.Storage.StorageFile])
$s=Aw ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$dec=Aw ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($s)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bmp=Aw ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$eng=[Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$res=Aw ($eng.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
Write-Output $res.Text
"""
    try:
        r = subprocess.run(["powershell","-NoProfile","-NonInteractive","-Command",ps],
                           capture_output=True, timeout=15)
        text = r.stdout.decode("utf-8", errors="ignore").strip()
        if text: return text
    except Exception:
        pass
    # Tesseract fallback
    try:
        import pytesseract
        from PIL import ImageEnhance
        for p in [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break
        img2 = img.convert("L")
        img2 = ImageEnhance.Contrast(img2).enhance(3.0)
        img2 = img2.point(lambda px: 0 if px < 160 else 255)
        return pytesseract.image_to_string(img2, config="--psm 6 -l eng")
    except Exception:
        return ""
    finally:
        try: os.unlink(path)
        except Exception: pass


def parse_ocr(text: str) -> dict:
    """Parse OCR text into fragment dict with set, slot, main, substats."""
    result = {"set": None, "slot": None, "main": None, "subs": []}
    lines  = [l.strip() for l in text.splitlines() if l.strip()]

    STAT_NORM = {
        r"critical\s*chance|crit\s*rate":    "Crit Rate%",
        r"critical\s*damage|crit\s*dmg":     "Crit DMG%",
        r"damage\s*over\s*time|dot":         "DoT%",
        r"extra\s*damage":                   "Extra DMG%",
        r"ego\s*recov\w*":                   "EGO Recovery",
        r"justice\s*damage":                 "Justice DMG%",
        r"order\s*damage":                   "Order DMG%",
        r"void\s*damage":                    "Void DMG%",
        r"instinct\s*damage":                "Instinct DMG%",
        r"passion\s*damage":                 "Passion DMG%",
        r"chaos\s*damage":                   "Chaos DMG%",
        r"\bdefense\b|\bdefence\b":          "DEF",
        r"\battack\b":                       "ATK",
        r"h[e3][a@][l1]th|\bhp\b|\bhealth\b": "HP",
    }

    def norm_stat(raw: str):
        r = raw.lower().strip()
        has_pct = "%" in r
        r = r.replace("%", "").strip()
        for pat, name in STAT_NORM.items():
            if re.search(pat, r):
                if has_pct and not name.endswith("%"):
                    return name + "%"
                return name
        return None

    def extract_val(rest: str):
        rest = rest.replace(",", ".")
        pct = re.findall(r'[+\-]?\d+\.?\d*\s*%', rest)
        if pct: return pct[-1].strip()
        flat = re.findall(r'[+\-]?\d+\.?\d*', rest)
        if flat: return flat[-1].strip()
        return None

    # Set + slot from title line
    for line in lines:
        t = line.lower()
        for alias, canonical in sorted(SET_ALIASES.items(), key=lambda x: -len(x[0])):
            if alias in t:
                result["set"] = canonical
                # Slot suffix
                words = t.split()
                for n in (1, 2):
                    suffix = " ".join(words[-n:]) if len(words) >= n else ""
                    from data import TITLE_SLOT_MAP
                    if suffix in TITLE_SLOT_MAP:
                        result["slot"] = TITLE_SLOT_MAP[suffix]
                break
        if result["set"]:
            break

    # Stats
    STAT_LINE = re.compile(
        r'(critical\s*(?:chance|damage|rate)|crit\s*(?:rate|dmg|damage)|'
        r'damage\s*over\s*time|extra\s*damage|ego\s*recov\w*|'
        r'justice\s*damage|order\s*damage|void\s*damage|instinct\s*damage|'
        r'passion\s*damage|chaos\s*damage|'
        r'defense|defence|attack|h[e3][a@][l1]th|health)(.*)',
        re.IGNORECASE
    )

    seen_names = set()
    entries = []
    for line in lines:
        m = STAT_LINE.search(line)
        if not m: continue
        name = norm_stat(m.group(1))
        if not name: continue
        rest = m.group(2)
        val  = extract_val(rest)
        if not val: continue
        # Skip tiny values that are upgrade counts
        try:
            fval = float(val.replace("%","").replace("+",""))
            if not "%" in val:
                if name == "HP" and fval < 7: continue
                if name == "DEF" and fval < 2.5: continue
                if name == "ATK" and fval < 4: continue
        except Exception:
            pass
        key = name + val
        if key not in seen_names:
            seen_names.add(key)
            entries.append({"name": name, "value": val})

    # First entry = main stat (orange in game, shown first)
    if entries:
        # Attribute DMG% is always main
        attr = next((e for e in entries if "DMG%" in e["name"] and
                     any(x in e["name"] for x in ["Justice","Order","Void","Instinct","Passion","Chaos"])), None)
        if attr:
            result["main"] = attr
            result["subs"] = [e for e in entries if e is not attr]
        elif result["slot"] in FIXED_MAIN:
            fixed_name = FIXED_MAIN[result["slot"]]
            match = next((e for e in entries if e["name"] == fixed_name), None)
            if match:
                result["main"] = match
                result["subs"] = [e for e in entries if e is not match]
            else:
                result["main"] = entries[0]
                result["subs"] = entries[1:]
        else:
            result["main"] = entries[0]
            result["subs"] = entries[1:]

    result["subs"] = result["subs"][:4]
    return result


# ── Settings ──────────────────────────────────────────────────────────────────
def load_settings():
    defaults = {"primary_char": "All", "show_below": 38}
    try:
        with open(SETTINGS_FILE) as f:
            return {**defaults, **json.load(f)}
    except Exception:
        return defaults

def save_settings(s):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(s, f, indent=2)
    except Exception:
        pass


# ── Overlay ───────────────────────────────────────────────────────────────────
class CZNOverlay:

    def __init__(self):
        self.settings = load_settings()
        self._last_result = None

        self.root = tk.Tk()
        self.root.title("CZN Fragment Rater")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.96)
        self.root.configure(bg=C["bg"])

        self._build_ui()

        sw, sh = pyautogui.size()
        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        x = max(0, sw - w - 20)
        y = max(0, sh // 2 - h // 2)
        self.root.geometry(f"+{x}+{y}")
        self.root.deiconify(); self.root.lift()

        # F9 hotkey
        self.root.bind_all("<F9>", lambda e: self._scan())

    # ── UI build ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Title / drag bar
        tb = tk.Frame(self.root, bg="#111827", pady=6, padx=10)
        tb.pack(fill="x")
        tk.Label(tb, text="CZN-Overlay", bg="#111827", fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(tb, text="F9 = scan", bg="#111827", fg=C["muted"],
                 font=("Segoe UI", 8)).pack(side="left", padx=8)
        tk.Button(tb, text="✕", command=self.root.destroy,
                  bg="#111827", fg=C["muted"], relief="flat",
                  font=("Segoe UI", 10), cursor="hand2", bd=0).pack(side="right")
        tb.bind("<ButtonPress-1>",   lambda e: (setattr(self,"_dx",e.x), setattr(self,"_dy",e.y)))
        tb.bind("<B1-Motion>", lambda e: self.root.geometry(
            f"+{self.root.winfo_x()+e.x-self._dx}+{self.root.winfo_y()+e.y-self._dy}"))
        self._dx = self._dy = 0

        # Two-column body
        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True)

        self._build_char_col(body)
        tk.Frame(body, bg=C["border"], width=1).pack(side="left", fill="y")
        self._build_input_col(body)

    # ── Left: character column ─────────────────────────────────────────────────
    def _build_char_col(self, parent):
        outer = tk.Frame(parent, bg=C["bg"], width=210)
        outer.pack(side="left", fill="y")
        outer.pack_propagate(False)

        # Fixed top area (Optimise for + primary char block)
        self._char_top = tk.Frame(outer, bg=C["bg"], padx=10, pady=8)
        self._char_top.pack(fill="x")

        # Scrollable list below
        list_frame = tk.Frame(outer, bg=C["bg"])
        list_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_frame, bg=C["bg"], highlightthickness=0, width=190)
        sb = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._char_list = tk.Frame(canvas, bg=C["bg"], padx=10, pady=4)
        win_id = canvas.create_window((0, 0), window=self._char_list, anchor="nw")

        def _on_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win_id, width=canvas.winfo_width())
        self._char_list.bind("<Configure>", _on_resize)

        def _on_wheel(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)

        # Keep reference to old _char_col for compatibility
        self._char_col = self._char_top
        self._refresh_chars()

    def _refresh_chars(self):
        for w in self._char_top.winfo_children():
            w.destroy()
        for w in self._char_list.winfo_children():
            w.destroy()

        top = self._char_top
        lst = self._char_list
        res     = self._last_result
        primary = self.settings["primary_char"]

        # ── Fixed top: selector ──
        tk.Label(top, text="Optimise for", bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        char_options = ["All"] + sorted(CHARS.keys())
        pv = tk.StringVar(value=primary)
        pcb = ttk.Combobox(top, textvariable=pv, values=char_options,
                           font=("Segoe UI", 10), width=16, state="readonly")
        pcb.pack(fill="x", pady=(2, 6))

        def _on_pchar(e=None):
            self.settings["primary_char"] = pv.get()
            save_settings(self.settings)
            if hasattr(self, "_set_cb"):
                char = pv.get()
                ns = self._sets_for(char) if char != "All" else SETS
                self._set_cb["values"] = ns
                if self._set_var.get() not in ns:
                    self._set_var.set(ns[0])
            self._refresh_chars()
        pcb.bind("<<ComboboxSelected>>", _on_pchar)

        # ── Fixed top: primary char block ──
        if primary != "All":
            pf = tk.Frame(top, bg="#1a1f28", padx=8, pady=8)
            pf.pack(fill="x", pady=(0, 4))
            prim_cs = next((cs for cs in res["char_scores"] if cs["char"] == primary), None) if res else None
            if prim_cs:
                h = tk.Frame(pf, bg="#1a1f28"); h.pack(fill="x")
                tk.Label(h, text=prim_cs["char"], bg="#1a1f28", fg=C["text"],
                         font=("Segoe UI",11,"bold")).pack(side="left")
                verd = res["verdict"]
                vbg = C["p1_bg"] if verd=="Keep" else C["p2_bg"] if verd=="Maybe" else C["bad_bg"]
                vfg = C["p1_fg"] if verd=="Keep" else C["p2_fg"] if verd=="Maybe" else C["bad_fg"]
                tk.Label(h, text=verd, bg=vbg, fg=vfg,
                         font=("Segoe UI",8,"bold"), padx=6).pack(side="right")
                sc = tk.Frame(pf, bg="#1a1f28"); sc.pack(fill="x", pady=(6,0))
                g2 = prim_cs["grade"] if prim_cs["is_maxed"] else prim_cs["best_grade"]
                tk.Label(sc, text=g2, bg=prim_cs["color"], fg=C["bg"],
                         font=("Segoe UI",22,"bold"), width=3, pady=2).pack(side="left", padx=(0,8))
                wf = tk.Frame(sc, bg="#1a1f28"); wf.pack(side="left")
                nf = tk.Frame(wf, bg="#1a1f28"); nf.pack(anchor="w")
                tk.Label(nf, text="Now", bg="#1a1f28", fg=C["muted"], font=("Segoe UI",8)).pack(side="left")
                tk.Label(nf, text=f"  {prim_cs['score']:.1f}", bg="#1a1f28", fg=res["color"],
                         font=("Segoe UI",13,"bold")).pack(side="left")
                if not prim_cs["is_maxed"]:
                    bf = tk.Frame(wf, bg="#1a1f28"); bf.pack(anchor="w")
                    tk.Label(bf, text="Best +5", bg="#1a1f28", fg=C["muted"], font=("Segoe UI",8)).pack(side="left")
                    bc = C["a"] if prim_cs["max_score"]>=43 else C["b"] if prim_cs["max_score"]>=30 else C["c"]
                    tk.Label(bf, text=f"  {prim_cs['max_score']:.1f}", bg="#1a1f28", fg=bc,
                             font=("Segoe UI",13,"bold")).pack(side="left")
            else:
                tk.Label(pf, text="Rate a fragment →", bg="#1a1f28", fg=C["muted"],
                         font=("Segoe UI",9)).pack()

        # ── Scrollable list ──
        lbl = "All characters" if primary == "All" else "Other characters"
        tk.Label(lst, text=lbl, bg=C["bg"], fg=C["muted"],
                 font=("Segoe UI",9)).pack(anchor="w", pady=(0,3))

        if res:
            all_cs = res["char_scores"]
            if primary != "All":
                all_cs = [cs for cs in all_cs if cs["char"] != primary]
            vis = [cs for cs in all_cs if (cs["max_score"] if not cs["is_maxed"] else cs["score"]) >= 18]
            hid = [cs for cs in all_cs if (cs["max_score"] if not cs["is_maxed"] else cs["score"]) < 18]
            last_rec = None
            for cs in vis:
                cur_rec = cs.get("recommended", True)
                if last_rec is True and not cur_rec:
                    tk.Frame(lst, bg=C["border"], height=1).pack(fill="x", pady=3)
                    tk.Label(lst, text="not recommended for this set",
                             bg=C["bg"], fg="#555", font=("Segoe UI",8)).pack(anchor="w", pady=(0,2))
                last_rec = cur_rec
                row = tk.Frame(lst, bg=C["bg"]); row.pack(fill="x", pady=1)
                tk.Label(row, text=cs["char"], bg=C["bg"],
                         fg=C["text"] if cur_rec else C["muted"],
                         font=("Segoe UI",10)).pack(side="left")
                ir = tk.Frame(row, bg=C["bg"]); ir.pack(side="right")
                g  = cs["grade"] if cs["is_maxed"] else cs["best_grade"]
                s_now = cs["score"]
                s_max = cs["max_score"]
                c2 = C["a"] if s_max>=43 else C["b"] if s_max>=30 else C["c"]
                tk.Label(ir, text=g, bg=cs["color"], fg=C["bg"],
                         font=("Segoe UI",8,"bold"), padx=4).pack(side="right", padx=(2,0))
                if cs["is_maxed"]:
                    # +5: only current score
                    tk.Label(ir, text=f"{s_now:.0f}", bg=C["bg"], fg=c2,
                             font=("Segoe UI",9,"bold")).pack(side="right", padx=(0,2))
                else:
                    # not maxed: show now → max
                    tk.Label(ir, text=f"{s_max:.0f}", bg=C["bg"], fg=c2,
                             font=("Segoe UI",9,"bold")).pack(side="right")
                    tk.Label(ir, text="→", bg=C["bg"], fg=C["muted"],
                             font=("Segoe UI",8)).pack(side="right", padx=1)
                    tk.Label(ir, text=f"{s_now:.0f}", bg=C["bg"], fg=C["muted"],
                             font=("Segoe UI",9)).pack(side="right", padx=(0,1))
            if hid:
                tk.Label(lst, text=f"+ {len(hid)} below C",
                         bg=C["bg"], fg=C["bad_fg"], font=("Segoe UI",8)).pack(anchor="w")
        else:
            tk.Label(lst, text="No data yet", bg=C["bg"], fg=C["muted"],
                     font=("Segoe UI",9)).pack(anchor="w")


    # ── Right: input column ────────────────────────────────────────────────────
    def _build_input_col(self, parent):
        pf = tk.Frame(parent, bg=C["bg"], padx=10, pady=10)
        pf.pack(side="left", fill="both", expand=True)

        # ── Scan button ──
        scan_row = tk.Frame(pf, bg=C["bg"]); scan_row.pack(fill="x", pady=(0,8))
        tk.Button(scan_row, text="📷  Scan (F9)", command=self._scan,
                  bg="#1a1f28", fg=C["text"], relief="flat",
                  font=("Segoe UI",9), cursor="hand2", padx=8, pady=4).pack(side="left")
        self._scan_status = tk.Label(scan_row, text="", bg=C["bg"], fg=C["muted"],
                                      font=("Segoe UI",8))
        self._scan_status.pack(side="left", padx=6)

        # ── Set + Slot ──
        r1 = tk.Frame(pf, bg=C["bg"]); r1.pack(fill="x", pady=(0,6))
        lf = tk.Frame(r1, bg=C["bg"]); lf.pack(side="left", fill="x", expand=True, padx=(0,6))
        tk.Label(lf, text="Set", bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w")
        _primary_now = self.settings.get("primary_char", "All")
        _init_sets   = self._sets_for(_primary_now) if _primary_now != "All" else SETS
        self._set_var = tk.StringVar(value=_init_sets[0])
        self._set_cb  = ttk.Combobox(lf, textvariable=self._set_var,
                                     values=_init_sets,
                                     font=("Segoe UI",10), width=20, state="readonly")
        self._set_cb.pack(fill="x")

        rf = tk.Frame(r1, bg=C["bg"]); rf.pack(side="left")
        tk.Label(rf, text="Slot", bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w")
        self._slot_var = tk.StringVar(value=SLOTS[0][1])
        self._slot_cb  = ttk.Combobox(rf, textvariable=self._slot_var,
                                      values=[s[1] for s in SLOTS],
                                      font=("Segoe UI",10), width=14, state="readonly")
        self._slot_cb.pack()
        self._slot_cb.bind("<<ComboboxSelected>>", self._on_slot_change)

        # ── Upgrade + Main ──
        r2 = tk.Frame(pf, bg=C["bg"]); r2.pack(fill="x", pady=(0,8))
        uf = tk.Frame(r2, bg=C["bg"]); uf.pack(side="left", padx=(0,8))
        tk.Label(uf, text="Upgrade", bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w")
        self._upg_var = tk.StringVar(value="+1")
        ttk.Combobox(uf, textvariable=self._upg_var, values=["+1","+2","+3","+4","+5"],
                     font=("Segoe UI",10), width=5, state="readonly").pack()

        mf = tk.Frame(r2, bg=C["bg"]); mf.pack(side="left", fill="x", expand=True)
        tk.Label(mf, text="Main stat", bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w")
        self._main_var = tk.StringVar(value=MAIN_BY_SLOT["shock"][0])
        self._main_cb  = ttk.Combobox(mf, textvariable=self._main_var,
                                      values=MAIN_BY_SLOT["shock"],
                                      font=("Segoe UI",10), width=14, state="readonly")
        self._main_cb.pack(fill="x")

        # ── Sub-stats ──
        tk.Frame(pf, bg=C["border"], height=1).pack(fill="x", pady=(0,6))
        tk.Label(pf, text="Sub-stats  (edit values if OCR is wrong)",
                 bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w", pady=(0,4))

        self._sub_vars = []
        for i in range(4):
            row = tk.Frame(pf, bg=C["bg"]); row.pack(fill="x", pady=2)
            tk.Label(row, text=str(i+1), bg=C["bg"], fg=C["muted"],
                     font=("Segoe UI",9), width=2).pack(side="left")
            sv = tk.StringVar(value="")
            scb = ttk.Combobox(row, textvariable=sv, values=ALL_SUBSTATS,
                               font=("Segoe UI",10), width=13, state="readonly")
            scb.pack(side="left", padx=(0,4))
            vv = tk.StringVar(value="")
            ve = tk.Entry(row, textvariable=vv, font=("Segoe UI",10),
                          width=7, bg="#1a1f28", fg=C["text"],
                          insertbackground=C["text"], relief="flat",
                          highlightthickness=1, highlightbackground=C["border"])
            ve.pack(side="left")
            self._sub_vars.append((sv, vv))

        # ── Rate button ──
        tk.Frame(pf, bg=C["border"], height=1).pack(fill="x", pady=(8,6))
        tk.Button(pf, text="Rate fragment", command=self._rate,
                  bg=C["p1_bg"], fg=C["p1_fg"], relief="flat",
                  font=("Segoe UI",10,"bold"), pady=6, cursor="hand2").pack(fill="x")

        # ── Sub-stat result ──
        self._result_frame = tk.Frame(pf, bg=C["bg"])
        self._result_frame.pack(fill="x", pady=(8,0))

    def _sets_for(self, char_name):
        rec  = list(CHARS[char_name]["sets"]) if char_name in CHARS else []
        rest = sorted(s for s in SETS if s not in rec)
        return rec + rest

    def _on_slot_change(self, e=None):
        key   = SLOT_LABEL_TO_KEY.get(self._slot_var.get(), "shock")
        mains = MAIN_BY_SLOT.get(key, ["ATK"])
        self._main_cb["values"] = mains
        if self._main_var.get() not in mains:
            self._main_var.set(mains[0])
        self._main_cb.configure(state="disabled" if key in FIXED_MAIN else "readonly")

    # ── Scan ──────────────────────────────────────────────────────────────────
    def _scan(self, e=None):
        self._scan_status.config(text="Scanning…", fg=C["muted"])
        self.root.update()
        # Hide overlay before screenshot
        self.root.withdraw()
        import time; time.sleep(0.15)
        try:
            sw, sh = pyautogui.size()
            img = ImageGrab.grab(bbox=(0, 0, sw, sh))
            text = _run_ocr(img)
            parsed = parse_ocr(text)
            self.root.after(0, lambda: self._fill_fields(parsed))
        except Exception as ex:
            self.root.after(0, lambda: self._scan_status.config(
                text=f"Error: {ex}", fg="#ff6b6b"))
        finally:
            self.root.after(0, self.root.deiconify)
            self.root.after(0, self.root.lift)

    def _fill_fields(self, parsed: dict):
        # Set
        if parsed.get("set") and parsed["set"] in SETS:
            self._set_var.set(parsed["set"])

        # Slot
        if parsed.get("slot"):
            label = next((v for k, v in SLOTS if k == parsed["slot"]), None)
            if label:
                self._slot_var.set(label)
                self._on_slot_change()

        # Main
        if parsed.get("main"):
            name = parsed["main"]["name"]
            if name in self._main_cb["values"]:
                self._main_var.set(name)

        # Subs
        for i, (sv, vv) in enumerate(self._sub_vars):
            if i < len(parsed.get("subs", [])):
                s = parsed["subs"][i]
                if s["name"] in ALL_SUBSTATS:
                    sv.set(s["name"])
                    vv.set(s["value"].replace("+","").replace("%",""))
            else:
                sv.set(""); vv.set("")

        self._scan_status.config(text="✓ Scanned — check and correct", fg=C["b"])

    # ── Rate ──────────────────────────────────────────────────────────────────
    def _rate(self):
        slot_key   = SLOT_LABEL_TO_KEY.get(self._slot_var.get(), "shock")
        upgrade    = int(self._upg_var.get().replace("+",""))
        main_name  = self._main_var.get()
        main_val   = {"name": main_name, "value": "—", "fval": 0.0}

        subs = []
        for sv, vv in self._sub_vars:
            name = sv.get()
            raw  = vv.get().strip().replace(",",".")
            if not name or not raw:
                continue
            try:
                fval = float(raw)
            except ValueError:
                continue
            subs.append({"name": name, "value": raw, "fval": fval})

        frag = {
            "set":       self._set_var.get(),
            "slot":      slot_key,
            "upgrade":   f"+{upgrade}",
            "main_stat": main_val,
            "sub_stats": subs,
        }

        self._last_result = score_fragment(frag)
        self._refresh_chars()
        self._show_quality()

    def _show_quality(self):
        for w in self._result_frame.winfo_children():
            w.destroy()
        res     = self._last_result
        primary = self.settings["primary_char"]
        if not res: return

        # ── Not recommended warning ──
        selected_set = self._set_var.get()
        prim_cs = None
        if primary != "All":
            prim_cs = next((cs for cs in res["char_scores"] if cs["char"] == primary), None)
            if prim_cs and not prim_cs.get("recommended", True):
                warn = tk.Frame(self._result_frame, bg="#2a1a00", padx=6, pady=4)
                warn.pack(fill="x", pady=(0,6))
                tk.Label(warn, text=f"⚠  {selected_set} is not recommended for {primary}",
                         bg="#2a1a00", fg="#ffaa44", font=("Segoe UI",8),
                         wraplength=260).pack(anchor="w")

        tk.Frame(self._result_frame, bg=C["border"], height=1).pack(fill="x", pady=(0,6))

        # ── Prydwen-style stat priority guide ──
        char_name = primary if primary != "All" else (res.get("best_char") or "")
        if char_name and char_name in CHARS:
            cd = CHARS[char_name]
            cw = cd["weights"]

            tk.Label(self._result_frame,
                     text=f"Stat priority for {char_name}",
                     bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w", pady=(0,4))

            def stars(w):
                if w >= 0.9: return "★★"
                if w >= 0.6: return "★"
                return "◆"

            def sep(w1, w2):
                if w1 - w2 < 0.1:  return " = "
                if w1 - w2 < 0.3:  return " > "
                return " >>> "

            ranked = sorted([(w, s) for s, w in cw.items() if w >= 0.3], reverse=True)

            if ranked:
                parts = []
                for i, (w, stat) in enumerate(ranked):
                    parts.append(f"{stars(w)} {stat}")
                    if i < len(ranked) - 1:
                        parts.append(sep(w, ranked[i+1][0]))

                priority_text = "".join(parts)
                tk.Label(self._result_frame, text=priority_text,
                         bg=C["bg"], fg=C["p1_fg"],
                         font=("Segoe UI", 9), wraplength=240,
                         justify="left").pack(anchor="w", pady=(0,6))

        # ── Actual sub-stat results ──
        if prim_cs and prim_cs.get("details"):
            tk.Label(self._result_frame, text="This fragment:",
                     bg=C["bg"], fg=C["muted"], font=("Segoe UI",9)).pack(anchor="w", pady=(0,3))

            for d in prim_cs["details"]:
                w = d["w"]
                if w >= 0.9:   lbl,lbg,lfg = "BEST", C["p1_bg"], C["p1_fg"]
                elif w >= 0.6: lbl,lbg,lfg = "GOOD", C["p2_bg"], C["p2_fg"]
                elif w >= 0.3: lbl,lbg,lfg = "OK",   "#1a2a1a",  "#9ecf9e"
                else:          lbl,lbg,lfg = "SKIP",  C["bad_bg"],C["bad_fg"]
                row2 = tk.Frame(self._result_frame, bg=C["bg"]); row2.pack(fill="x", pady=1)
                tk.Label(row2, text=lbl, bg=lbg, fg=lfg,
                         font=("Segoe UI",8,"bold"), width=5, padx=2).pack(side="left")
                tk.Label(row2, text=f"  {d['name']}", bg=C["bg"], fg=C["text"],
                         font=("Segoe UI",10)).pack(side="left")
                tk.Label(row2, text=d["value"], bg=C["bg"], fg=C["muted"],
                         font=("Segoe UI",10)).pack(side="right")

            n_good   = prim_cs.get("n_good_subs", 0)
            is_maxed = prim_cs.get("is_maxed", False)
            upg      = int(str(res.get("upgrade","1")).replace("+",""))
            expl = f"{n_good}/4 useful stats · final" if is_maxed \
                   else f"{n_good}/4 useful stats · {5-upg} rolls left"
            tk.Label(self._result_frame, text=expl, bg=C["bg"], fg=C["bad_fg"],
                     font=("Segoe UI",8)).pack(anchor="w", pady=(4,0))

    def run(self):
        print("CZN Fragment Rater v5 — F9=scan | close X to quit")
        self.root.mainloop()


if __name__ == "__main__":
    CZNOverlay().run()
