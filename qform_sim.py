# -*- coding: utf-8 -*-
"""
QForm-Simulation aus dem Stichplan: Konfiguration erzeugen, Dateien schreiben, Lauf starten
==========================================================================================
Der eigentliche Aufbau in QForm steckt in qform_runner.py (läuft als eigener Prozess, damit eine stundenlange
Berechnung die Oberfläche nicht blockiert und auch nach dem Schließen des Rechners weiterläuft).
"""
import datetime
import glob
import json
import os
import re
import subprocess
import sys

import qform_runner

SHAPE_NAMES = {"rund": "Rund", "vier": "Vierkant", "acht": "Achtkant"}
SHAPE_SHORT = {"rund": "Ø", "vier": "VK", "acht": "SW"}

DEFAULT_TOOL = (r"C:\Users\QForm_user\Desktop\QForm_API_Stichplan_ReAlForm\00_resources\Tools"
                r"\Flachsattel_Pahnke.qshape")
DEFAULT_WP_MATERIAL = (r"C:\Users\QForm_user\Desktop\QForm_API_Stichplan_ReAlForm\00_resources\Workpiece"
                       r"\AA 6082_Leiber_WUMSI_DRX.qdat")
DEFAULT_QFORM_DIR = r"C:\QForm UK\12.0.5"
# Bekannte Sättel: Breite in Stabachse [mm] (aus der Bounding-Box in QForm)
KNOWN_DIES = {"flachsattel_pahnke.qshape": (300.0, 25.0)}     # Breite in Stabachse, Kantenradius [mm]
GEOMETRY_TYPES = [("Geometrie", "*.qshape *.step *.stp *.igs *.iges *.stl *.csv3d *.x_t *.sat"), ("Alle Dateien", "*.*")]
QDAT_TYPES = [("QForm-Werkstoff", "*.qdat"), ("Alle Dateien", "*.*")]


def _version_key(name):
    return tuple(int(x) for x in re.findall(r"\d+", name)) or (0,)


def find_qform_dirs():
    """QForm-Installationen mit Python-API (neueste zuerst)."""
    roots = [r"C:\QForm UK", r"C:\Program Files\QForm UK", r"C:\Program Files\QForm", r"C:\QForm",
             os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "QForm UK")]
    found = []
    for root in roots:
        for d in glob.glob(os.path.join(root, "*")):
            if os.path.isfile(os.path.join(d, "API", "App", "Python", "QFormAPI.py")) and \
                    os.path.isfile(os.path.join(d, "x64", "QFormCon.exe")):
                if d not in found:
                    found.append(d)
    return sorted(found, key=lambda d: _version_key(os.path.basename(d)), reverse=True)


def qform_dir_ok(d):
    return bool(d) and os.path.isfile(os.path.join(d, "API", "App", "Python", "QFormAPI.py"))


def is_db_path(s):
    s = (s or "").strip()
    return bool(re.match(r"^(std|doc|user|prj|qform):", s))


def die_width_known(path):
    """(Breite, ebene Arbeitsfläche, Kantenradius) eines bekannten Sattels oder None."""
    d = KNOWN_DIES.get(re.split(r"[\\/]", path or "")[-1].lower())
    return None if d is None else (d[0], d[0] - 2.0 * d[1], d[1])


def pass_name(k, shape, d, reheat=None):
    dim = f"{d:.0f}" if abs(d - round(d)) < 0.05 else f"{d:.1f}".replace(".", ",")
    name = f"Stich {k + 1} {SHAPE_SHORT[shape]}{' ' if shape != 'vier' else ''}{dim}"
    if reheat is not None:
        name += f" (Wiedererwärmung {reheat:g} °C)"
    return name


def sim_feed(row, bissfolge, f_greif=0.5):
    """Vorschub je Biss mit Vorzeichen. Die Längung eines Bisses s·(λ−1) verteilt sich auf beide Seiten: der Anteil
    f_greif verschiebt die Greifseite (Manipulator gibt nach), der Rest das freie Ende. Damit der nächste Biss genau
    an den vorigen anschließt:
      vom Manipulator weg:  +s·(1 + (λ−1)·(1 − f))   (f = 0: starrer Manipulator, Vorschub s·λ wie im Stichplan)
      zum Manipulator hin:  −s·(1 + (λ−1)·f)         (f = 0: Vorschub s wie Stichplan V1/V3)"""
    s = float(row["sB"])
    lam = float(row["feed"]) / s if s > 0 else 1.0
    f = min(1.0, max(0.0, float(f_greif)))
    if bissfolge == "manipulator":
        return s * (1.0 + (lam - 1.0) * (1.0 - f))       # Werkstück fährt nach +x
    return -s * (1.0 + (lam - 1.0) * f)                  # Werkstück fährt nach −x


def sim_blows(row, k, proc, bissfolge, reheat, extra_bites=0, f_greif=0.5):
    """Hubtabelle eines Stichs für die Simulation: (CoolingInAir, CoolingInTool, Feed, Rotation) je Hub.
    Wie der QForm-Export des Stichplans; Vorschub mit Vorzeichen der Bissfolge (sim_feed), Transportzeit auch nach
    einer Wiedererwärmung, keine Rückfahrt (jeder Stich wird in QForm neu positioniert)."""
    n_b = int(row["n_b"]) + int(extra_bites)
    n_rot = int(row["n_rot"])
    u = sim_feed(row, bissfolge, f_greif)
    air0 = proc["t_transfer"] if (k == 0 or reheat is not None) else proc["t_pause"]
    out = []
    for j in range(n_b):
        for q in range(n_rot):
            if j == 0 and q == 0:
                out.append((round(air0, 3), round(proc["t_aux"], 3), 0.0, 0.0))
            else:
                out.append((0.0, round(proc["t_aux"], 3), round(u, 1) if q == 0 else 0.0, round(row["rot_step"], 3)))
    return out


def build_config(plan, s, version):
    """plan: Stichplan aus dem Rechner (geo, res, proc, mat, block, passes); s: Einstellungen des Reiters.
    Rückgabe: (Konfiguration, Hinweise[(art, text)], Vorschau-Zeilen)."""
    notes = []
    res, geo, proc = plan["res"], plan["geo"], plan["proc"]
    rows = res["rows"]
    b = plan["block"]
    T_wp = proc["T_start"] if s["wp_T_mode"] == "plan" else s["wp_T"]
    T_tool = proc["T_die"] if s["tool_T_mode"] == "plan" else s["tool_T"]
    stiche, preview = [], []
    for k, r in enumerate(rows):
        g = geo[k]
        reheat = s["reheat"][k] if k < len(s["reheat"]) and k > 0 else None
        f_greif = s.get("elong_grip", 0.5)
        blows = sim_blows(r, k, proc, s["bissfolge"], reheat, s["extra_bites"], f_greif)
        name = pass_name(k, g["shape"], g["d"], reheat)
        stiche.append(dict(nr=k + 1, name=name, form=g["shape"], mass=round(g["d"], 3), laenge=round(g["L"], 1),
                           bisslaenge=round(r["sB"], 3), letzter_biss=round(r["s_last"], 3),
                           drehlagen=int(r["n_rot"]), bisse=int(r["n_b"]) + int(s["extra_bites"]),
                           vorschub=round(abs(sim_feed(r, s["bissfolge"], f_greif)), 2),
                           streckgrad=round(g["ratio"], 4), waermen=reheat, f_plan_MN=round(r["F"], 3),
                           T_kern_plan=round(r["T_core"], 1), blows=[list(x) for x in blows]))
        preview.append(dict(nr=k + 1, shape=g["shape"], d=g["d"], L=g["L"], blows=len(blows),
                            n_b=int(r["n_b"]) + int(s["extra_bites"]), n_rot=int(r["n_rot"]), rot=r["rot_step"],
                            feed=blows[int(r["n_rot"])][2] if len(blows) > int(r["n_rot"]) else 0.0,
                            air=blows[0][0], reheat=reheat))
    wp_geo = dict(quelle=s["wp_geo"], datei=s["wp_file"], form=b["shape"], d=b["d"], L=b["L"], r=b.get("r") or 0.0,
                  methode=s.get("wp_method", "auto"))
    lub = s["lubricant"]
    cfg = dict(
        format=qform_runner.FORMAT, version=1, programmversion=version,
        erstellt=datetime.datetime.now().isoformat(timespec="seconds"),
        qform_dir=s["qform_dir"], ausgabe_dir=s["out_dir"], projektname=s["name"],
        prozessname=f"Stichplan {plan['mat']['name']}",
        werkstueck=dict(bezeichnung=plan["mat"]["name"], geometrie=wp_geo, werkstoff=s["wp_mat"], temperatur=T_wp,
                        mikrostruktur=s["micro"], korngroesse_um=s["grain"]),
        werkzeug=dict(geometrie=s["tool_file"], arbeitsflaeche=s["tool_face"], werkstoff=s["tool_mat"],
                      temperatur=T_tool, temperatur_vererben=s["tool_inherit"], schmierstoff=lub, spalt=s["gap"]),
        presse=dict(typ="hydraulisch", nennkraft_MN=s["press_F"], v_mm_s=proc["v"]),
        manipulator=dict(grifflaenge=s["grip"], rueckstellung=s["give_back"], laengung_greifseite=s.get("elong_grip", 0.5),
                         widerstand_axial=s["res_ax"], widerstand_quer=s["res_cross"], aktiv_im_hub=False),
        bissfolge=s["bissfolge"], zusatzbisse=s["extra_bites"], sattelbreite_stichplan=proc["b_die"],
        max_rechenzeit_h=s["max_h"], qform_am_ende_schliessen=s["close"],
        stichplan=dict(rohblock=dict(form=b["shape"], d=b["d"], L=b["L"], r=b.get("r") or 0.0),
                       t_transport=proc["t_transfer"], t_neben=proc["t_aux"], t_pause=proc["t_pause"],
                       v=proc["v"], verweilzeit=proc["t_dwell"], sattelbreite=proc["b_die"],
                       bisslaengenmodus=proc.get("bite_mode")),
        stiche=stiche)

    # ---------------- Hinweise
    if not qform_dir_ok(s["qform_dir"]):
        notes.append(("err", f"QFormAPI.py nicht gefunden im QForm-Verzeichnis „{s['qform_dir']}“ "
                             f"(erwartet: API\\App\\Python\\QFormAPI.py)."))
    if not s["out_dir"]:
        notes.append(("err", "Ausgabeordner angeben."))
    if not os.path.isfile(s["tool_file"] or ""):
        notes.append(("err", f"Sattelgeometrie nicht gefunden: {s['tool_file'] or '(leer)'}"))
    if s["wp_geo"] == "datei" and not os.path.isfile(s["wp_file"] or ""):
        notes.append(("err", f"Werkstückgeometrie nicht gefunden: {s['wp_file'] or '(leer)'}"))
    if s["wp_geo"] == "datei":
        notes.append(("info", "Die Werkstückgeometrie aus der Datei muss zum Rohblock des Stichplans passen "
                              f"({SHAPE_NAMES[b['shape']]} {b['d']:g} × {b['L']:g} mm); sonst stimmen Längen und "
                              "Bisse nicht überein."))
    for what, val in (("Werkstückwerkstoff", s["wp_mat"]), ("Werkzeugwerkstoff", s["tool_mat"])):
        if not val:
            notes.append(("err", f"{what} angeben (.qdat-Datei oder Datenbankpfad)."))
        elif not is_db_path(val) and not os.path.isfile(val):
            notes.append(("err", f"{what}: Datei nicht gefunden: {val}"))
    if not lub or not is_db_path(lub):
        notes.append(("err", "Schmierstoff (Reibung) als Datenbankpfad angeben, z. B. std:lubricant/…"))
    w = die_width_known(s["tool_file"])
    if w and proc["b_die"] > w[1] + 0.5:
        notes.append(("warn", f"Sattelbreite im Stichplan {proc['b_die']:g} mm, die ebene Arbeitsfläche des Sattels ist "
                              f"nur {w[1]:g} mm breit ({w[0]:g} mm mit Kantenradien R {w[2]:g}). Längere Bisse lassen "
                              f"zwischen den Bissen ungeschmiedete Bereiche; Sattelbreite im Reiter Stichplan auf "
                              f"{w[1]:g} mm setzen."))
    if s["extra_bites"] and s["bissfolge"] == "manipulator":
        notes.append(("warn", "Zusatzbisse bei der Bissfolge vom Manipulator weg schieben das Werkstück über den Sattel "
                              "hinaus (Hübe ohne Kontakt). Sinnvoll sind sie bei der Bissfolge zum Manipulator hin, "
                              "dort schmieden sie das gegriffene Ende nach (wie Stichplan V3)."))
    if b["shape"] == "acht" and s["wp_geo"] == "rohblock":
        notes.append(("info", "Achtkant-Rohblock: Werkstück wird als Extrusion eines DXF-Querschnitts erzeugt."))
    if s["micro"]:
        notes.append(("info", "Kornentwicklung nur mit einem Werkstoff, der ein Gefügemodell enthält (z. B. JMAK/DRX)."))
    n_blows = sum(len(x["blows"]) for x in stiche)
    notes.append(("info", f"{len(stiche)} Operationen mit zusammen {n_blows} Hüben; Presse {s['press_F']:g} MN "
                          f"hydraulisch, {proc['v']:g} mm/s aus dem Stichplan; Stoppbedingung je Stich: Abstand der "
                          "Sättel = Maß des Stichs."))
    return cfg, notes, preview


# ======================================================================== Dateien
BAT = """@echo off
rem {titel}
chcp 65001 >nul
cd /d "%~dp0"
title {titel}
set "PYEXE="
set "PYARG="
where py >nul 2>nul && (set "PYEXE=py" & set "PYARG=-3")
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\\Programs\\Python\\Python3*") do if exist "%%D\\python.exe" set "PYEXE=%%D\\python.exe"
if not defined PYEXE set "PYEXE=python"
"%PYEXE%" %PYARG% -X utf8 "%~dp0qform_runner.py" "%~dp0simulation.json" {args}
echo.
echo Fertig - Protokoll im Ordner Protokolle
pause
"""

LIESMICH = """QForm-Simulation des Stichplans (Freiformschmieden-Rechner {version})
=====================================================================
Erstellt am {stand}. Werkstoff: {werkstoff}.

Start (vorher andere QForm-Fenster schließen, Lizenzplatz)
  Simulation_berechnen.bat      baut alle Stiche auf und rechnet sie nacheinander
  Simulation_Aufbau_pruefen.bat baut nur Stich 1 auf (keine Berechnung), QForm bleibt zur Kontrolle offen
  Benötigt Python 3 (python.org) und QForm UK mit API: {qform_dir}

Dateien
  simulation.json   alle Einstellungen und je Stich die Hubtabelle (Blows) aus dem Stichplan
  qform_runner.py   Aufbau und Berechnung über die QForm-Python-API
  Protokolle\\       Protokoll (*.log), Status und Ergebnis je Stich (*_ergebnis.json)
  *.qform           erzeugtes QForm-Projekt (Name mit Zeitstempel)

Stiche
{stiche}

Aufbau
  Sättel        {sattel}, Untersattel fest, Obersattel hydraulische Presse {presse:g} MN / {v:g} mm/s,
                Werkstoff {wz_werkstoff}, {T_wz:g} °C, Schmierstoff {lub}
  Werkstück     {wp_geo}, Werkstoff {wp_werkstoff}, {T_wp:g} °C
  Bissfolge     {bissfolge}
  Hübe          je Biss alle Drehlagen, vor jedem Hub {t_neben:g} s auf dem Untersattel (CoolingInTool),
                vor Hub 1 Transport {t_transport:g} s bzw. Pause {t_pause:g} s an Luft (CoolingInAir),
                Stoppbedingung: Abstand der Sättel = Maß des Stichs
"""


def runner_source():
    """Quelltext von qform_runner.py (auch im gepackten Programm, dort als Datendatei mitgeliefert)."""
    cands = [os.path.join(getattr(sys, "_MEIPASS", ""), "qform_runner.py"),
             os.path.join(os.path.dirname(os.path.abspath(qform_runner.__file__)), "qform_runner.py")]
    for c in cands:
        if c and os.path.isfile(c):
            with open(c, "r", encoding="utf-8") as fh:
                return fh.read()
    raise FileNotFoundError("qform_runner.py nicht gefunden")


def write_package(folder, cfg):
    """Konfiguration, Ablaufskript, Startdateien und Liesmich in den Ordner schreiben. Rückgabe: Dateiliste."""
    os.makedirs(folder, exist_ok=True)
    files = []
    p = os.path.join(folder, "simulation.json")
    with open(p, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=1)
    files.append(p)
    p = os.path.join(folder, "qform_runner.py")
    with open(p, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(runner_source())
    files.append(p)
    for name, titel, args in (("Simulation_berechnen.bat", "QForm-Simulation berechnen", ""),
                              ("Simulation_Aufbau_pruefen.bat", "QForm-Simulation Aufbau pruefen", "--nur-aufbau")):
        p = os.path.join(folder, name)
        with open(p, "w", encoding="utf-8", newline="\r\n") as fh:
            fh.write(BAT.format(titel=titel, args=args))
        files.append(p)
    st = []
    for x in cfg["stiche"]:
        st.append(f"  {x['nr']:>2}  {SHAPE_NAMES[x['form']]:<9}{x['mass']:>7g} mm  Länge {x['laenge']:>6.0f} mm  "
                  f"{len(x['blows']):>3} Hübe = {x['bisse']} Bisse × {x['drehlagen']} Drehlagen, Vorschub "
                  f"{x['vorschub']:.1f} mm" + (f"  <- Wiedererwärmung {x['waermen']:g} °C" if x["waermen"] is not None
                                               else ""))
    g = cfg["werkstueck"]["geometrie"]
    wp_geo = (f"aus dem Rohblock erzeugt ({SHAPE_NAMES[g['form']]} {g['d']:g} × {g['L']:g} mm)"
              if g["quelle"] == "rohblock" else os.path.basename(g["datei"]))
    p = os.path.join(folder, "LIESMICH.txt")
    with open(p, "w", encoding="utf-8-sig", newline="\r\n") as fh:
        fh.write(LIESMICH.format(
            version=cfg["programmversion"], stand=cfg["erstellt"].replace("T", " "),
            werkstoff=cfg["werkstueck"]["bezeichnung"], qform_dir=cfg["qform_dir"], stiche="\n".join(st),
            sattel=os.path.basename(cfg["werkzeug"]["geometrie"]), presse=cfg["presse"]["nennkraft_MN"],
            v=cfg["presse"]["v_mm_s"], wz_werkstoff=cfg["werkzeug"]["werkstoff"], T_wz=cfg["werkzeug"]["temperatur"],
            lub=cfg["werkzeug"]["schmierstoff"], wp_geo=wp_geo, wp_werkstoff=cfg["werkstueck"]["werkstoff"],
            T_wp=cfg["werkstueck"]["temperatur"],
            bissfolge=(("vom Manipulator weg, Vorschub = s·(1 + (λ−1)·(1 − f))" if cfg["bissfolge"] == "manipulator"
                        else "zum Manipulator hin, Vorschub = s·(1 + (λ−1)·f)")
                       + f", Anteil der Längung zur Greifseite f = {cfg['manipulator'].get('laengung_greifseite', 0.5):g}"),
            t_neben=cfg["stichplan"]["t_neben"], t_transport=cfg["stichplan"]["t_transport"],
            t_pause=cfg["stichplan"]["t_pause"]))
    files.append(p)
    return files


def start_run(cfg_path, nur_aufbau, log_path, status_path, stop_path):
    """Lauf als eigenen Prozess starten (gepackt: dieses Programm mit --qform-run, sonst Python + qform_runner.py)."""
    if getattr(sys, "frozen", False):
        args = [sys.executable, "--qform-run", cfg_path]
    else:
        args = [sys.executable, os.path.join(os.path.dirname(os.path.abspath(__file__)), "qform_runner.py"), cfg_path]
    if nur_aufbau:
        args.append("--nur-aufbau")
    args += ["--log", log_path, "--status", status_path, "--stop", stop_path]
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"          # neue, unabhängige Instanz des gepackten Programms
    kw = {}
    if sys.platform == "win32":
        kw["creationflags"] = 0x08000000                # CREATE_NO_WINDOW
    return subprocess.Popen(args, cwd=os.path.dirname(cfg_path), env=env, stdin=subprocess.DEVNULL,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, **kw)
