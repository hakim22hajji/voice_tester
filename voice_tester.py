"""
═══════════════════════════════════════════════════════════════
  EDGE TTS VOICE TESTER
  Test voices sentence-by-sentence.
  Compare RAW vs FFmpeg-Enhanced audio side by side.

  Install:  pip install edge-tts
  Optional: ffmpeg on PATH for audio enhancement
  No extra packages needed for playback (uses Windows MCI).
═══════════════════════════════════════════════════════════════
"""

import asyncio
import queue
import random
import re
import os
import threading
import tempfile
import subprocess
import shutil
import ctypes
from pathlib import Path
import tkinter as tk
from tkinter import ttk

import edge_tts

# ─────────────────────────────────────────────
#  COLOURS
# ─────────────────────────────────────────────

C = {
    "bg":       "#1a1a2e",
    "surface":  "#16213e",
    "surface2": "#0f3460",
    "accent":   "#e94560",
    "accent2":  "#533483",
    "green":    "#4ade80",
    "yellow":   "#facc15",
    "red":      "#f87171",
    "text":     "#eaeaea",
    "dim":      "#8892a4",
    "border":   "#2d3748",
}

# ─────────────────────────────────────────────
#  VOICES
# ─────────────────────────────────────────────

VOICES = [
    "en-GB-RyanNeural",
    "en-GB-SoniaNeural",
    "en-GB-LibbyNeural",
    "en-US-ChristopherNeural",
    "en-US-JennyNeural",
    "en-US-GuyNeural",
    "en-US-AriaNeural",
    "en-US-EricNeural",
    "fr-FR-RemyNeural",
    "fr-FR-DeniseNeural",
    "fr-FR-HenriNeural",
    "de-DE-ConradNeural",
    "es-ES-AlvaroNeural",
    "it-IT-DiegoNeural",
]

# ─────────────────────────────────────────────
#  SMART PROSODY  (subtle, paragraph-level)
# ─────────────────────────────────────────────

# Adjustments are relative offsets added to the user's base rate/pitch.
# Kept deliberately small so transitions feel natural, not theatrical.
SMART_PROSODY = {
    "dramatic": {
        "keywords": ["suddenly", "gasped", "terror", "horror", "darkness", "dead",
                     "dying", "blood", "fear", "trembled", "shock", "stunned", "collapsed"],
        "rate": -8, "pitch": -3, "vol": -3,
    },
    "action": {
        "keywords": ["ran", "sprinted", "rushed", "exploded", "crashed", "burst",
                     "attacked", "fought", "chased", "escaped", "leapt", "fired", "shot"],
        "rate": +6, "pitch": +2, "vol": +3,
    },
    "tender": {
        "keywords": ["whispered", "gently", "softly", "love", "embraced", "smiled",
                     "tears", "quiet", "peaceful", "warm", "tender", "murmured", "sighed"],
        "rate": -10, "pitch": +2, "vol": -5,
    },
    "suspense": {
        "keywords": ["slowly", "crept", "shadow", "listened", "waited", "watching",
                     "silence", "hidden", "lurking", "uneasy", "dread", "approached"],
        "rate": -8, "pitch": -2, "vol": -4,
    },
    "joyful": {
        "keywords": ["laughed", "celebrated", "joyful", "delighted", "wonderful",
                     "amazing", "fantastic", "excited", "thrilled", "happy", "dancing"],
        "rate": +4, "pitch": +4, "vol": +3,
    },
    "sad": {
        "keywords": ["wept", "cried", "sobbed", "mourned", "grief", "sorrow",
                     "alone", "lonely", "despair", "hopeless", "broken", "farewell"],
        "rate": -12, "pitch": -5, "vol": -7,
    },
}


def _parse_rate(s: str) -> int:
    return int(s.replace("%", "").replace("+", "") or "0")


def _parse_pitch(s: str) -> int:
    return int(s.replace("Hz", "").replace("+", "") or "0")


def detect_prosody(text: str, base_rate: str, base_pitch: str) -> dict:
    """
    Score each emotion against the paragraph text, pick the winner,
    apply small offsets on top of the user's base values, add tiny jitter.
    Returns dict with keys: rate, pitch, vol, emotion.
    """
    tl = text.lower()
    scores = {
        name: sum(1 for kw in data["keywords"] if kw in tl)
        for name, data in SMART_PROSODY.items()
    }
    best = max(scores, key=scores.get)
    adj  = SMART_PROSODY[best] if scores[best] > 0 else {"rate": 0, "pitch": 0, "vol": 0}
    label = best if scores[best] > 0 else "neutral"

    r = _parse_rate(base_rate)  + adj["rate"]  + random.randint(-2, 2)
    p = _parse_pitch(base_pitch) + adj["pitch"] + random.randint(-1, 1)
    v = adj["vol"]

    r = max(-50, min(50, r))
    p = max(-20, min(20, p))
    v = max(-20, min(20, v))

    return {
        "rate":    f"{r:+d}%",
        "pitch":   f"{p:+d}Hz",
        "vol":     f"{v:+d}%",
        "emotion": label,
    }


def split_paragraphs(text: str) -> list[str]:
    """Split on blank lines; fall back to single sentences if no breaks."""
    paras = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
    if len(paras) <= 1:
        # No paragraph breaks — split by sentence so the list is still useful
        paras = [s.strip() for s in re.split(r'(?<=[.!?])\s+', text.strip()) if s.strip()]
    return paras


# ─────────────────────────────────────────────
#  WINDOWS MCI AUDIO PLAYER  (no extra deps)
# ─────────────────────────────────────────────

_mci = ctypes.windll.winmm.mciSendStringW


def _mci_cmd(cmd: str) -> int:
    return _mci(cmd, None, 0, None)


class AudioPlayer:
    """
    Plays an MP3 file using Windows MCI (built-in, no extra packages).
    Each instance uses a unique alias so multiple players work in parallel.
    """

    _counter = 0

    def __init__(self):
        AudioPlayer._counter += 1
        self._alias = f"tts_player_{AudioPlayer._counter}"
        self._open = False

    def play(self, path: str) -> tuple[bool, str]:
        self.stop()
        path = str(Path(path).resolve())
        err = _mci_cmd(f'open "{path}" type mpegvideo alias {self._alias}')
        if err:
            return False, f"MCI open error {err}"
        err = _mci_cmd(f"play {self._alias}")
        if err:
            _mci_cmd(f"close {self._alias}")
            return False, f"MCI play error {err}"
        self._open = True
        return True, "Playing…"

    def stop(self):
        if self._open:
            _mci_cmd(f"stop {self._alias}")
            _mci_cmd(f"close {self._alias}")
            self._open = False

    def is_playing(self) -> bool:
        if not self._open:
            return False
        buf = ctypes.create_unicode_buffer(128)
        _mci(f"status {self._alias} mode", buf, 128, None)
        return buf.value == "playing"

    def __del__(self):
        try:
            self.stop()
        except Exception:
            pass


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def check_ffmpeg() -> bool:
    try:
        r = subprocess.run(["ffmpeg", "-version"], capture_output=True, timeout=3)
        return r.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False



async def _tts_async(text: str, voice: str, rate: str, pitch: str, path: str, volume: str = "+0%"):
    comm = edge_tts.Communicate(text, voice, rate=rate, pitch=pitch, volume=volume)
    audio = bytearray()
    async for chunk in comm.stream():
        if chunk["type"] == "audio":
            audio.extend(chunk["data"])
    if not audio:
        raise RuntimeError(
            "TTS returned no audio — check your internet connection "
            f"and that the voice '{voice}' exists."
        )
    with open(path, "wb") as f:
        f.write(bytes(audio))



def enhance_file(raw: str, out: str) -> tuple[bool, str]:
    cmd = [
        "ffmpeg", "-y", "-i", raw,
        "-af",
        (
            "equalizer=f=3000:width_type=o:width=2:g=2,"
            "equalizer=f=200:width_type=o:width=2:g=-1.5,"
            "acompressor=threshold=-20dB:ratio=3:attack=5:release=80:makeup=2,"
            "loudnorm=I=-16:TP=-1.5:LRA=11"
        ),
        "-b:a", "192k", "-loglevel", "error", out,
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=30)
        if r.returncode == 0:
            return True, "Enhanced"
        return False, f"FFmpeg error: {r.stderr.decode()[:120]}"
    except FileNotFoundError:
        return False, "FFmpeg not found"
    except subprocess.TimeoutExpired:
        return False, "FFmpeg timed out"


# ─────────────────────────────────────────────
#  AUDIO PANEL WIDGET
# ─────────────────────────────────────────────

class AudioPanel(tk.Frame):
    """One half of the split comparison view."""

    def __init__(self, parent, title: str, header_color: str, **kw):
        super().__init__(parent, bg=C["surface"], **kw)
        self._player = AudioPlayer()
        self._path: str | None = None

        # Header bar
        hdr = tk.Frame(self, bg=header_color)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text=title,
            bg=header_color, fg=C["text"],
            font=("Segoe UI", 12, "bold"),
            pady=12,
        ).pack()

        # Status
        self._status_var = tk.StringVar(value="Idle")
        self._status_lbl = tk.Label(
            self, textvariable=self._status_var,
            bg=C["surface"], fg=C["dim"],
            font=("Segoe UI", 9),
        )
        self._status_lbl.pack(pady=(16, 2))

        # File info
        self._info_var = tk.StringVar(value="—")
        tk.Label(
            self, textvariable=self._info_var,
            bg=C["surface"], fg=C["text"],
            font=("Segoe UI", 10),
        ).pack(pady=4)

        # Indeterminate progress bar (shown while generating)
        self._bar = ttk.Progressbar(self, mode="indeterminate", length=220)
        self._bar.pack(pady=8)

        # Buttons
        row = tk.Frame(self, bg=C["surface"])
        row.pack(pady=12)
        self._play_btn = tk.Button(
            row, text="▶  Play", width=10,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            command=self.play,
        )
        self._play_btn.pack(side="left", padx=4)
        tk.Button(
            row, text="■  Stop", width=8,
            bg=C["surface2"], fg=C["text"],
            font=("Segoe UI", 10),
            relief="flat", cursor="hand2",
            command=self.stop,
        ).pack(side="left", padx=4)

        self._set_play_enabled(False)
        self._poll()

    # ── Public API ────────────────────────────

    def set_loading(self, msg: str = "Generating…"):
        self._path = None
        self._set_play_enabled(False)
        self._set_status(msg, C["yellow"])
        self._info_var.set("—")
        self._bar.start(12)

    def set_ready(self, path: str, label: str = ""):
        self._path = path
        self._bar.stop()
        kb = os.path.getsize(path) / 1024 if os.path.exists(path) else 0
        info = f"{kb:.1f} KB"
        if label:
            info += f"  ·  {label}"
        self._info_var.set(info)
        self._set_status("Ready — click Play to listen", C["green"])
        self._set_play_enabled(True)

    def set_ready_and_play(self, path: str, label: str = ""):
        """Atomically set the file ready and immediately start playing."""
        self.set_ready(path, label)
        self.play()

    def set_error(self, msg: str):
        self._path = None
        self._bar.stop()
        self._set_play_enabled(False)
        self._set_status(f"Error: {msg}", C["red"])
        self._info_var.set("—")

    def set_warning(self, msg: str):
        self._bar.stop()
        self._set_status(msg, C["yellow"])

    def clear(self):
        self._player.stop()
        self._path = None
        self._bar.stop()
        self._set_play_enabled(False)
        self._set_status("Idle — click a sentence above", C["dim"])
        self._info_var.set("—")

    def play(self):
        if not self._path:
            self._set_status("No audio loaded — click a sentence first", C["yellow"])
            return
        ok, msg = self._player.play(self._path)
        if not ok:
            self._set_status(f"Playback error: {msg}", C["red"])
            from tkinter import messagebox
            messagebox.showerror("Playback Error", msg)
        else:
            self._set_status("Playing…", C["green"])

    def stop(self):
        self._player.stop()
        self._set_status("Stopped", C["dim"])

    # ── Internal ──────────────────────────────

    def _set_play_enabled(self, enabled: bool):
        """Toggle play button appearance AND clickability together."""
        if enabled:
            self._play_btn.config(
                state="normal",
                bg=C["accent"], fg="white",
                cursor="hand2",
            )
        else:
            self._play_btn.config(
                state="disabled",
                bg=C["border"], fg=C["dim"],
                cursor="arrow",
            )

    def _set_status(self, msg: str, color: str):
        self._status_var.set(msg)
        self._status_lbl.config(fg=color)

    def _poll(self):
        """Update Playing… / Done label in real time."""
        if self._player.is_playing():
            if self._status_var.get() != "Playing…":
                self._set_status("Playing…", C["green"])
        elif self._path and self._status_var.get() == "Playing…":
            self._set_status("Done ✓", C["green"])
        self.after(300, self._poll)


# ─────────────────────────────────────────────
#  MAIN APPLICATION
# ─────────────────────────────────────────────

class VoiceTesterApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Edge TTS Voice Tester")
        self.geometry("1060x720")
        self.minsize(840, 580)
        self.configure(bg=C["bg"])

        self._tmpdir = tempfile.mkdtemp(prefix="tts_tester_")
        self._items: list[dict] = []   # [{text, rate, pitch, vol, emotion}]
        self._selected_idx: int | None = None
        self._play_all_stop = threading.Event()
        self._worker: threading.Thread | None = None
        self._ffmpeg_ok = check_ffmpeg()
        self._ui_queue: queue.Queue = queue.Queue()

        self._setup_styles()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._drain_ui_queue()

    # ── Styles ────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame",    background=C["bg"])
        s.configure("TLabel",    background=C["bg"], foreground=C["text"], font=("Segoe UI", 10))
        s.configure("TCombobox",
                    fieldbackground=C["surface2"], background=C["surface2"],
                    foreground=C["text"], selectbackground=C["accent"])
        s.configure("Vertical.TScrollbar",
                    background=C["surface2"], troughcolor=C["surface"], borderwidth=0)
        s.configure("TProgressbar",
                    troughcolor=C["surface"], background=C["accent2"])

    # ── UI ────────────────────────────────────

    def _build_ui(self):

        # ══ TOP BAR ══════════════════════════════════════════════
        top = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        top.pack(fill="x")

        title_row = tk.Frame(top, bg=C["bg"])
        title_row.pack(fill="x", pady=(0, 8))
        tk.Label(
            title_row, text="Edge TTS Voice Tester",
            bg=C["bg"], fg=C["accent"],
            font=("Segoe UI", 17, "bold"),
        ).pack(side="left")

        # Badges (right-aligned)
        ffmpeg_bg   = C["green"]  if self._ffmpeg_ok else C["yellow"]
        ffmpeg_txt  = " FFmpeg ✓ " if self._ffmpeg_ok else " FFmpeg ✗ "
        tk.Label(title_row, text=ffmpeg_txt,
                 bg=ffmpeg_bg, fg=C["bg"],
                 font=("Segoe UI", 9, "bold"),
                 padx=4).pack(side="right", padx=(4, 0))
        tk.Label(title_row, text=" Windows MCI ✓ ",
                 bg=C["green"], fg=C["bg"],
                 font=("Segoe UI", 9, "bold"),
                 padx=4).pack(side="right", padx=(4, 0))

        # Text input
        tk.Label(top, text="Text to test:", bg=C["bg"], fg=C["dim"],
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._textbox = tk.Text(
            top, height=4, wrap="word",
            bg=C["surface"], fg=C["text"],
            insertbackground=C["text"],
            relief="flat", font=("Segoe UI", 10),
            padx=8, pady=6,
        )
        self._textbox.pack(fill="x", pady=(2, 8))
        self._textbox.insert("1.0",
            "Hello! This is a first test sentence. "
            "Here comes a slightly longer second sentence. "
            "And a third one to round things out nicely.")
        self._textbox.bind("<Control-Return>", lambda _e: self._parse_and_load())

        # Controls row
        ctrl = tk.Frame(top, bg=C["bg"])
        ctrl.pack(fill="x")

        tk.Label(ctrl, text="Voice:", bg=C["bg"], fg=C["text"]).pack(side="left")
        self._voice_var = tk.StringVar(value="en-GB-RyanNeural")
        ttk.Combobox(ctrl, textvariable=self._voice_var,
                     values=VOICES, width=26, state="readonly"
                     ).pack(side="left", padx=(4, 18))

        tk.Label(ctrl, text="Rate:", bg=C["bg"], fg=C["text"]).pack(side="left")
        self._rate_var = tk.StringVar(value="-5%")
        tk.Entry(ctrl, textvariable=self._rate_var, width=6,
                 bg=C["surface2"], fg=C["text"], insertbackground=C["text"],
                 relief="flat").pack(side="left", padx=(4, 18))

        tk.Label(ctrl, text="Pitch:", bg=C["bg"], fg=C["text"]).pack(side="left")
        self._pitch_var = tk.StringVar(value="+0Hz")
        tk.Entry(ctrl, textvariable=self._pitch_var, width=7,
                 bg=C["surface2"], fg=C["text"], insertbackground=C["text"],
                 relief="flat").pack(side="left", padx=(4, 18))

        # Smart Prosody toggle
        self._smart_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            ctrl, text="Smart Prosody",
            variable=self._smart_var,
            bg=C["bg"], fg=C["text"],
            selectcolor=C["surface2"],
            activebackground=C["bg"], activeforeground=C["text"],
            font=("Segoe UI", 10),
        ).pack(side="left", padx=(0, 12))

        tk.Button(
            ctrl, text="  Parse  [Ctrl+Enter]  ",
            bg=C["accent"], fg="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            command=self._parse_and_load,
        ).pack(side="left", padx=(0, 6))

        self._play_all_btn = tk.Button(
            ctrl, text="▶  Play All",
            bg=C["surface2"], fg=C["text"],
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2",
            command=self._toggle_play_all,
        )
        self._play_all_btn.pack(side="left")

        # ══ PARAGRAPH LIST ═══════════════════════════════════════
        mid = tk.Frame(self, bg=C["bg"], padx=16)
        mid.pack(fill="x")

        self._list_label_var = tk.StringVar(value="Paragraphs  —  click one to generate & compare")
        tk.Label(
            mid, textvariable=self._list_label_var,
            bg=C["bg"], fg=C["dim"], font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(8, 2))

        list_wrap = tk.Frame(mid, bg=C["border"], bd=1, relief="solid")
        list_wrap.pack(fill="x")
        vsb = ttk.Scrollbar(list_wrap, orient="vertical")
        self._listbox = tk.Listbox(
            list_wrap, yscrollcommand=vsb.set,
            bg=C["surface"], fg=C["text"],
            selectbackground=C["accent"], selectforeground="white",
            font=("Segoe UI", 10), relief="flat",
            height=5, activestyle="none", cursor="hand2",
        )
        vsb.config(command=self._listbox.yview)
        vsb.pack(side="right", fill="y")
        self._listbox.pack(side="left", fill="both", expand=True)
        self._listbox.bind("<<ListboxSelect>>", self._on_sentence_click)

        # ══ SPLIT PANELS ═════════════════════════════════════════
        panels = tk.Frame(self, bg=C["bg"], padx=16, pady=12)
        panels.pack(fill="both", expand=True)
        panels.columnconfigure(0, weight=1)
        panels.columnconfigure(1, weight=1)
        panels.rowconfigure(0, weight=1)

        self._raw_panel = AudioPanel(
            panels,
            title="RAW  —  Edge TTS, no post-processing",
            header_color=C["surface2"],
        )
        self._raw_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        self._enh_panel = AudioPanel(
            panels,
            title="ENHANCED  —  FFmpeg EQ + compression + loudnorm",
            header_color=C["accent2"],
        )
        self._enh_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        if not self._ffmpeg_ok:
            self._enh_panel.set_warning(
                "FFmpeg not found — install FFmpeg and add it to\n"
                "your PATH to enable audio enhancement"
            )

        # ══ STATUS BAR ═══════════════════════════════════════════
        self._status_var = tk.StringVar(value="Ready.  Enter some text and click Parse.")
        tk.Label(
            self, textvariable=self._status_var,
            bg=C["surface"], fg=C["dim"],
            font=("Segoe UI", 9), anchor="w",
            padx=12, pady=5,
        ).pack(fill="x", side="bottom")

    # ── Logic ─────────────────────────────────

    def _parse_and_load(self):
        text = self._textbox.get("1.0", "end").strip()
        if not text:
            return

        base_rate  = self._rate_var.get().strip()  or "-5%"
        base_pitch = self._pitch_var.get().strip() or "+0Hz"
        use_smart  = self._smart_var.get()

        paras = split_paragraphs(text)

        self._items = []
        self._listbox.delete(0, "end")
        for i, para in enumerate(paras, 1):
            if use_smart:
                p = detect_prosody(para, base_rate, base_pitch)
            else:
                p = {"rate": base_rate, "pitch": base_pitch,
                     "vol": "+0%", "emotion": "manual"}
            self._items.append({"text": para, **p})

            emotion_tag = f"[{p['emotion']}] " if use_smart else ""
            preview     = para if len(para) <= 80 else para[:77] + "…"
            self._listbox.insert("end", f"  {i}.  {emotion_tag}{preview}")

        unit = "paragraph" if '\n' in text else "sentence"
        plural = "s" if len(paras) != 1 else ""
        self._list_label_var.set(
            f"{len(paras)} {unit}{plural}  —  click one to generate & compare"
        )
        self._raw_panel.clear()
        self._enh_panel.clear()
        self._selected_idx = None
        self._status(f"{len(paras)} {unit}{plural} parsed.  Click one to generate audio.")

    def _on_sentence_click(self, _event=None):
        sel = self._listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == self._selected_idx:
            self._raw_panel.play()
            if self._ffmpeg_ok:
                self._enh_panel.play()
            return
        self._selected_idx = idx
        self._generate(idx)

    def _drain_ui_queue(self):
        """Run all pending UI callbacks posted from worker threads. Runs in main thread."""
        while True:
            try:
                fn = self._ui_queue.get_nowait()
                fn()
            except queue.Empty:
                break
        self.after(50, self._drain_ui_queue)

    def _post(self, fn):
        """Thread-safe: schedule fn() to run on the main thread."""
        self._ui_queue.put(fn)

    def _generate(self, idx: int):
        if self._worker and self._worker.is_alive():
            self._status("Still generating — please wait a moment…")
            return

        item   = self._items[idx]
        text   = item["text"]
        voice  = self._voice_var.get()
        rate   = item["rate"]
        pitch  = item["pitch"]
        vol    = item["vol"]
        emotion = item["emotion"]
        ffmpeg = self._ffmpeg_ok
        total  = len(self._items)

        raw_path = os.path.join(self._tmpdir, f"s{idx}_raw.mp3")
        enh_path = os.path.join(self._tmpdir, f"s{idx}_enh.mp3")

        self._raw_panel.set_loading("Generating TTS…")
        self._enh_panel.set_loading("Waiting for raw audio…")
        self._status(f"Generating {idx + 1}/{total}  [{emotion}  {rate}  {pitch}]…")

        # Build the label shown in the panel once ready
        raw_label = f"{emotion}  ·  rate {rate}  ·  pitch {pitch}"

        def worker():
            # 1 — TTS  (pass volume via edge-tts volume param)
            try:
                asyncio.run(_tts_async(text, voice, rate, pitch, raw_path, vol))
            except Exception as exc:
                err = str(exc)
                self._post(lambda: self._raw_panel.set_error(err))
                self._post(lambda: self._enh_panel.set_error("Skipped (TTS failed)"))
                self._post(lambda: self._status(f"TTS error: {err}"))
                return

            self._post(lambda: self._raw_panel.set_ready_and_play(raw_path, raw_label))

            # 2 — FFmpeg enhancement
            if not ffmpeg:
                self._post(lambda: self._enh_panel.set_warning(
                    "FFmpeg not available — install it to enable enhancement"
                ))
                self._post(lambda: self._status(f"Item {idx + 1} ready (raw only)."))
                return

            self._post(lambda: self._enh_panel.set_loading("Enhancing with FFmpeg…"))
            ok, msg = enhance_file(raw_path, enh_path)
            if ok:
                enh_label = f"enhanced  ·  {raw_label}"
                self._post(lambda: self._enh_panel.set_ready_and_play(enh_path, enh_label))
                self._post(lambda: self._status(f"Item {idx + 1} ready."))
            else:
                self._post(lambda: self._enh_panel.set_error(msg))
                self._post(lambda: self._status(f"Enhancement failed: {msg}"))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _highlight_item(self, idx: int):
        """Highlight a row in the listbox (main thread only)."""
        self._listbox.selection_clear(0, "end")
        self._listbox.selection_set(idx)
        self._listbox.see(idx)
        self._selected_idx = idx

    def _toggle_play_all(self):
        """Start Play All (both panels), or stop it if already running."""
        if self._worker and self._worker.is_alive():
            self._play_all_stop.set()
            return

        if not self._items:
            self._status("Parse text first, then click Play All.")
            return

        self._play_all_stop.clear()
        self._play_all_btn.config(text="■  Stop", bg=C["red"])
        self._raw_panel.set_loading("Generating all paragraphs…")
        self._enh_panel.set_loading("Waiting for raw audio…")

        voice  = self._voice_var.get()
        ffmpeg = self._ffmpeg_ok
        items  = list(self._items)

        def worker():
            import time

            def _concat(paths: list, out: str) -> bool:
                """Concatenate MP3 files using FFmpeg. Returns True on success."""
                if len(paths) == 1:
                    import shutil as _sh
                    _sh.copy2(paths[0], out)
                    return True
                list_file = out + ".txt"
                with open(list_file, "w", encoding="utf-8") as f:
                    for p in paths:
                        f.write(f"file '{p.replace(chr(92), '/')}'\n")
                try:
                    r = subprocess.run(
                        ["ffmpeg", "-y", "-f", "concat", "-safe", "0",
                         "-i", list_file, "-c", "copy", out],
                        capture_output=True, timeout=120,
                    )
                    return r.returncode == 0
                except Exception:
                    return False

            # ── Phase 1: generate all raw paragraphs ───────────
            raw_paths = []
            for i, item in enumerate(items):
                if self._play_all_stop.is_set():
                    break
                self._post(lambda i=i: self._highlight_item(i))
                self._post(lambda i=i: self._status(
                    f"Generating {i+1}/{len(items)}  [{items[i]['emotion']}]…"
                ))
                # Include voice name in filename so changing voice invalidates cache
                voice_tag = re.sub(r'[^\w]', '_', voice)
                raw_path = os.path.join(self._tmpdir, f"s{i}_{voice_tag}_raw.mp3")
                if not (os.path.exists(raw_path) and os.path.getsize(raw_path) > 0):
                    try:
                        asyncio.run(_tts_async(
                            item["text"], voice,
                            item["rate"], item["pitch"], raw_path, item["vol"]
                        ))
                    except Exception as exc:
                        self._post(lambda e=str(exc): self._status(f"TTS error: {e}"))
                        self._post(lambda: self._raw_panel.set_error("TTS failed"))
                        self._post(lambda: self._enh_panel.set_error("Skipped"))
                        self._post(lambda: self._reset_play_all_btn())
                        return
                raw_paths.append(raw_path)

            if self._play_all_stop.is_set() or not raw_paths:
                self._post(lambda: self._status("Stopped."))
                self._post(lambda: self._raw_panel.clear())
                self._post(lambda: self._enh_panel.clear())
                self._post(lambda: self._reset_play_all_btn())
                return

            # ── Phase 2: combine raw → one file ─────────────────
            all_raw = os.path.join(self._tmpdir, "all_raw.mp3")
            all_enh = os.path.join(self._tmpdir, "all_enhanced.mp3")

            if ffmpeg:
                self._post(lambda: self._status("Combining raw audio…"))
                raw_ok = _concat(raw_paths, all_raw)
            else:
                raw_ok = False  # will fall back to sequential below

            # ── Phase 3: enhance combined raw ───────────────────
            enh_ok = False
            if ffmpeg and raw_ok and not self._play_all_stop.is_set():
                self._post(lambda: self._enh_panel.set_loading("Enhancing full audio…"))
                self._post(lambda: self._status("Enhancing with FFmpeg…"))
                ok, msg = enhance_file(all_raw, all_enh)
                enh_ok = ok
                if not ok:
                    self._post(lambda m=msg: self._enh_panel.set_error(m))

            # ── Phase 4: load both panels ────────────────────────
            n = len(raw_paths)
            raw_label = f"all {n} paragraph(s) · raw"
            enh_label = f"all {n} paragraph(s) · enhanced"

            if raw_ok:
                self._post(lambda: self._raw_panel.set_ready(all_raw, raw_label))
            if enh_ok:
                self._post(lambda: self._enh_panel.set_ready(all_enh, enh_label))

            # ── Phase 5: play ────────────────────────────────────
            if raw_ok:
                # Panels are loaded — auto-play raw. User clicks Enhanced ▶ to compare.
                # Use the panel's own player (set_ready_and_play) — no second player here.
                self._post(lambda: self._raw_panel.set_ready_and_play(all_raw, raw_label))
                self._post(lambda: self._status(
                    "Playing full text (raw). Click ▶ Play in the Enhanced panel to compare."
                ))
            else:
                # No FFmpeg — sequential playback using a dedicated player
                # (keeps panel player free for the Stop button)
                self._post(lambda: self._status("Playing all paragraphs sequentially…"))
                seq_player = AudioPlayer()
                for i, path in enumerate(raw_paths):
                    if self._play_all_stop.is_set():
                        break
                    self._post(lambda i=i: self._highlight_item(i))
                    self._post(lambda i=i: self._status(
                        f"Playing {i+1}/{len(raw_paths)}  [{items[i]['emotion']}]…"
                    ))
                    ok, _ = seq_player.play(path)
                    if ok:
                        while seq_player.is_playing() and not self._play_all_stop.is_set():
                            time.sleep(0.15)
                seq_player.stop()

            self._post(lambda: self._status(
                "Stopped." if self._play_all_stop.is_set() else
                "Done. Click ▶ Play in either panel to replay and compare."
            ))
            self._post(lambda: self._reset_play_all_btn())
            self._post(lambda: self._status(
                "Stopped." if self._play_all_stop.is_set() else "Done playing all."
            ))
            self._post(lambda: self._reset_play_all_btn())

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _reset_play_all_btn(self):
        self._play_all_btn.config(text="▶  Play All", bg=C["surface2"])

    def _status(self, msg: str):
        self._status_var.set(msg)

    def _on_close(self):
        shutil.rmtree(self._tmpdir, ignore_errors=True)
        self.destroy()


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = VoiceTesterApp()
    app.mainloop()
