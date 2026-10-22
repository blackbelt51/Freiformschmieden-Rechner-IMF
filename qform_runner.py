# -*- coding: utf-8 -*-
"""
QForm-Simulation des Stichplans aufbauen und rechnen (QForm UK 12, Python-API)
=============================================================================
Wird vom Freiformschmieden-Rechner (Reiter „QForm-Simulation“) gestartet oder eigenständig:

    python qform_runner.py simulation.json               alle Stiche aufbauen und Stich für Stich rechnen
    python qform_runner.py simulation.json --nur-aufbau  nur Stich 1 aufbauen und speichern, QForm bleibt offen

Die Konfiguration (simulation.json) schreibt der Freiformschmieden-Rechner: Geometrien, Werkstoffe, Temperaturen,
Wiedererwärmung, Schmierstoff (Reibung), Presse und je Stich die Hubtabelle (Blows) aus dem Stichplan.

Aufbau im globalen Koordinatensystem (wie Stichplan V1/V3 des ReAlForm-Projekts):
  z  Pressrichtung; Arbeitsfläche des Untersattels bei z = 0, Obersattel darüber
  x  Werkstückachse = Bissrichtung; Sattel mittig bei x = 0, Sattelbreite in x
  Manipulator greift das hintere Werkstückende (+x), das vordere Ende (−x) ist frei.
  Bissfolge „vom Manipulator weg“: erster Biss am gegriffenen Ende (ragt eine Bisslänge unter den Sattel),
                                   Vorschub +u (Werkstück fährt nach +x)
  Bissfolge „zum Manipulator hin“: erster Biss am freien Ende (ragt eine Bisslänge unter den Sattel),
                                   Vorschub −u (Werkstück fährt nach −x)
  In beiden Fällen wird die ganze Länge geschmiedet wie im Stichplan, auch die Greifzone des Manipulators.
Jeder Stich ist eine eigene Operation; ab Stich 2 werden Werkstück und Sättel geerbt und das Werkstück an seiner
Ist-Geometrie für den ersten Biss neu positioniert. Nur Standardbibliothek und QFormAPI.py aus dem QForm-Verzeichnis.
"""
import base64
import datetime
import io  # noqa: F401  (von QFormAPI benötigt; im gepackten Programm sonst nicht enthalten)
import json
import math
import os
import struct
import subprocess  # noqa: F401
import sys
import time
import traceback
import typing  # noqa: F401
import urllib.parse  # noqa: F401
from pathlib import Path

FORMAT = "Freiformschmieden-Rechner QForm-Simulation"

# Vorlage hydraulischer Antrieb (QForm-Binärformat wie project_export_as_script); hyd_speed [m/s] und
# hyd_max_load [N] werden überschrieben. Aus dem erprobten Stichplan-V3-Skript übernommen.
_ANTRIEB_VORLAGE = (
    "+XZvYmoKCgH7AwoFAfsI+gGYA2RzY5sA+gGYBXBhcnRz+gKYCURpcmVjdGlvbvoBmAV1c2FnZf36AZgFdmFsdWX6A5gF"
    "UG9pbnT6AZgBeJQAAAAAAAAAAPoBmAF5lAAAAAAAAAAA+gGYAXqUAAAAAAAA8L/6AZgJdGltZXN0YW1wh6APTxMotdwB"
    "+gGYA3ZlcgH6AZgJaHlkX3NwZWVklHsUrkfhepQ/+gGYDGh5ZF9tYXhfbG9hZJQAAAAA0BJjQfoBmAhoeWRfdHlwZZgD"
    "aW5k+gGYD2h5ZF9sb2FkcGVyY2VudJSamZmZmZnpPw==")

ANTRIEB_FEST = "qform/114/+OZ"            # QForm: fester Antrieb (Untersattel)
ENDE = {"Done", "Error", "Canceled", "HaspKeyNotFound"}


class Abbruch(Exception):
    pass


# ======================================================================== Hilfen ohne QForm
def antrieb_bytes(kraft_N, v_mm_s):
    """Antriebsdatei (hydraulische Presse) aus der Vorlage mit Nennkraft und Geschwindigkeit."""
    raw = bytearray(base64.b64decode(_ANTRIEB_VORLAGE))
    for schluessel, zahl in [(b"hyd_speed", v_mm_s / 1000.0), (b"hyd_max_load", kraft_N)]:
        i = raw.find(schluessel + b"\x94")
        if i < 0:
            raise RuntimeError(f"Antriebsvorlage: Feld {schluessel!r} nicht gefunden")
        p = i + len(schluessel) + 1
        raw[p:p + 8] = struct.pack("<d", float(zahl))
    return bytes(raw)


def querschnitt_dxf(form, d, r=0.0):
    """DXF (R12) mit dem Querschnitt in der xy-Ebene, Mittelpunkt im Ursprung, Flächen parallel zu x und y.
    form: 'rund' (d = Durchmesser), 'vier' (d = Kantenlänge), 'acht' (d = Schlüsselweite); r = Kantenradius."""
    ent = []

    def line(p, q_):
        ent.append(("LINE", [(10, p[0]), (20, p[1]), (30, 0.0), (11, q_[0]), (21, q_[1]), (31, 0.0)]))

    def arc(c, rad, a0, a1):
        ent.append(("ARC", [(10, c[0]), (20, c[1]), (30, 0.0), (40, rad), (50, a0), (51, a1)]))

    if form == "rund":
        ent.append(("CIRCLE", [(10, 0.0), (20, 0.0), (30, 0.0), (40, d / 2.0)]))
    else:
        n = 4 if form == "vier" else 8
        a = d / 2.0
        r = max(0.0, min(float(r or 0.0), a * 0.999))
        half = math.pi / n
        normals = [2 * math.pi * j / n for j in range(n)]
        pts = []                                # je Ecke: (Tangentenpunkt Seite j, Tangentenpunkt Seite j+1, Zentrum)
        for j in range(n):
            th = normals[j] + half
            dc = (a - r) / math.cos(half)
            c = (dc * math.cos(th), dc * math.sin(th))
            p0 = (c[0] + r * math.cos(normals[j]), c[1] + r * math.sin(normals[j]))
            p1 = (c[0] + r * math.cos(normals[j] + 2 * half), c[1] + r * math.sin(normals[j] + 2 * half))
            pts.append((p0, p1, c))
        for j in range(n):
            p0, p1, c = pts[j]
            if r > 1e-9:
                arc(c, r, math.degrees(normals[j]), math.degrees(normals[j] + 2 * half))
            nxt = pts[(j + 1) % n][0]
            line(p1, nxt)
    out = io.StringIO()
    w = out.write
    w("0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1009\n9\n$INSUNITS\n70\n4\n0\nENDSEC\n")
    w("0\nSECTION\n2\nENTITIES\n")
    for typ, codes in ent:
        w(f"0\n{typ}\n8\n0\n")
        for code, val in codes:
            w(f"{code}\n{val:.6f}\n")
    w("0\nENDSEC\n0\nEOF\n")
    return out.getvalue()


def lade_konfiguration(pfad):
    with open(pfad, "r", encoding="utf-8-sig") as fh:
        cfg = json.load(fh)
    if cfg.get("format") != FORMAT:
        raise ValueError(f"{pfad} ist keine Simulationskonfiguration des Freiformschmieden-Rechners.")
    return cfg


# ======================================================================== Ablauf mit QForm
class Lauf:
    def __init__(self, cfg, nur_aufbau=False, log_pfad=None, status_pfad=None, stop_pfad=None):
        self.cfg = cfg
        self.nur_aufbau = nur_aufbau
        self.zeit = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.aus = Path(cfg["ausgabe_dir"])
        self.name = cfg.get("projektname") or "Stichplan"
        self.projekt = self.aus / f"{self.name}_{self.zeit}.qform"
        prot = self.aus / "Protokolle"
        self.log_pfad = Path(log_pfad) if log_pfad else prot / f"{self.name}_{self.zeit}.log"
        self.status_pfad = Path(status_pfad) if status_pfad else prot / f"{self.name}_{self.zeit}_status.json"
        self.erg_pfad = self.log_pfad.with_name(self.log_pfad.stem + "_ergebnis.json")
        self.stop_pfad = Path(stop_pfad) if stop_pfad else self.aus / "STOP"
        self.QF = None
        self.q = None
        self.logdatei = None
        self.ergebnisse = []
        self.projekt_angelegt = False
        self.qform_bereit = False
        self.status = dict(zustand="startet", stich=0, stiche=len(cfg["stiche"]), hub=0, huebe=0, prozesszeit=0.0,
                           meldung="", projekt=str(self.projekt), log=str(self.log_pfad), start=self.zeit,
                           nur_aufbau=nur_aufbau, ergebnisse=self.ergebnisse)

    # ------------------------------------------------------------------ Protokoll und Status
    def log(self, text="", tiefe=0):
        z = f"[{datetime.datetime.now():%H:%M:%S}] " + "    " * tiefe + str(text)
        if sys.stdout is not None:
            try:
                print(z, flush=True)
            except Exception:
                pass
        if self.logdatei:
            self.logdatei.write(z + "\n")
            self.logdatei.flush()

    def melde(self, **kw):
        self.status.update(kw)
        self.status["aktualisiert"] = datetime.datetime.now().isoformat(timespec="seconds")
        try:
            tmp = self.status_pfad.with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(self.status, fh, ensure_ascii=False, indent=1)
            os.replace(tmp, self.status_pfad)
        except OSError:
            pass

    def stop_pruefen(self):
        if self.stop_pfad.exists():
            raise Abbruch("Abbruch angefordert")

    # ------------------------------------------------------------------ QForm-Grundfunktionen
    def ev(self, x):
        return getattr(x, "value", str(x))

    def wert(self, typ, oid, pfad, w):
        q = self.q
        self.QF.property_set(q.Property(typ, oid, pfad, q.PropertyType.Value, str(w)))

    def objekt(self, typ, oid, pfad, ziel_typ, ziel_id):
        self.QF.property_set_object(self.q.ObjectIdProperty(typ, oid, pfad, ziel_typ, ziel_id))

    def box(self, typ, oid):
        b = Box(self.QF.mesh_bound_box(self.q.MeshObjectId(typ, oid, 0)))
        if not b.gueltig():
            raise RuntimeError(f"ungültige Bounding-Box für {self.ev(typ)} {oid}: {b}")
        return b

    def knoten(self):
        m = self.QF.mesh_nodes_get(self.q.MeshObjectId(self.WP, 1, 0))
        return list(zip(m.x, m.y, m.z))

    def verschieben(self, typ, oid, dx, dy, dz):
        if max(abs(dx), abs(dy), abs(dz)) > 1e-6:
            self.QF.object_move(self.q.ObjectMove(typ, oid, dx, dy, dz, False))

    def drehen(self, typ, oid, punkt, achse, winkel):
        self.QF.object_rotate(self.q.ObjectRotate(typ, oid, punkt[0], punkt[1], punkt[2],
                                                  achse[0], achse[1], achse[2], winkel, False))

    # ------------------------------------------------------------------ Start
    def qform_starten(self):
        qdir = self.cfg["qform_dir"]
        api = os.path.join(qdir, "API", "App", "Python")
        if not os.path.isfile(os.path.join(api, "QFormAPI.py")):
            raise FileNotFoundError(f"QFormAPI.py nicht gefunden in {api}")
        sys.dont_write_bytecode = True
        if api not in sys.path:
            sys.path.insert(0, api)
        import QFormAPI as q  # noqa: E402
        self.q = q
        self.WP, self.TOOL, self.OP = q.ObjectType.Workpiece, q.ObjectType.Tool, q.ObjectType.Operation
        # eigener Instanzname je Lauf, damit parallele Läufe sich nicht dieselbe QForm-Instanz teilen
        self._umgebung_fuer_qform()
        self.QF = q.QForm(self.cfg.get("instanz") or f"Freiformschmieden-Rechner_{self.zeit}")
        self.QF.qform_dir_set(os.path.join(qdir, "x64"))
        self.QF.qform_start()
        v = self.QF.qform_version()
        t0 = time.time()
        lic = self.QF.license_key_status()
        while not lic.key_is_ok and time.time() - t0 < float(self.cfg.get("lizenz_wartezeit_s", 30)):
            time.sleep(2)
            lic = self.QF.license_key_status()
        self.log(f"QForm {v.v1}.{v.v2}.{v.v3}.{v.v4} gestartet, Lizenz ok={lic.key_is_ok} '{lic.key_status_msg}'")
        if not lic.key_is_ok:
            raise RuntimeError("Kein Lizenzschlüssel frei. Andere QForm-Fenster schließen und erneut starten.")
        self.QF.print(q.TraceMsg(f"Stichplan aus dem Freiformschmieden-Rechner ({self.zeit})", q.MessageType.info))
        # API-Lizenz (Feature „API initial data access“) wird vom Lizenzserver teils verzögert vergeben
        t0 = time.time()
        while True:
            try:
                self.QF.project_new()
                self.qform_bereit = True
                break
            except Exception as ex:
                if "License required" not in str(ex) or time.time() - t0 > float(self.cfg.get("lizenz_wartezeit_s", 60)):
                    if "License required" in str(ex):
                        raise RuntimeError(f"QForm-Lizenz für die API nicht verfügbar ({ex}). In QForm unter "
                                           "Help > License prüfen, ob ‚API initial data access‘ frei ist, und andere "
                                           "QForm-Fenster schließen.") from ex
                    raise
                time.sleep(3)
        if time.time() - t0 > 1:
            self.log(f"API-Lizenz nach {time.time() - t0:.0f} s erhalten", 1)

    @staticmethod
    def _umgebung_fuer_qform():
        """Im gepackten Programm: QForm nicht mit dem DLL-Suchpfad und den Pfaden des Entpackordners (_MEI…) starten,
        sonst hält das weiterlaufende QForm Dateien darin fest und der Ordner lässt sich beim Beenden nicht löschen."""
        mei = getattr(sys, "_MEIPASS", None)
        if not mei:
            return
        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.kernel32.SetDllDirectoryW(None)
            except Exception:
                pass
        mei_n = os.path.normcase(os.path.abspath(mei))
        for k in list(os.environ):
            v = os.environ.get(k, "")
            if k.upper() == "PATH":
                teile = [t for t in v.split(os.pathsep) if t and not os.path.normcase(os.path.abspath(t)).startswith(mei_n)]
                os.environ[k] = os.pathsep.join(teile)
            elif v and os.path.normcase(v).startswith(mei_n):
                os.environ.pop(k, None)

    # ------------------------------------------------------------------ Datenbankobjekte
    def db_objekt(self, quelle, art, standard=None):
        """Werkstoff aus .qdat-Datei importieren (doc:<art>/<Dateiname>) oder Datenbankpfad direkt verwenden."""
        quelle = (quelle or "").strip() or (standard or "")
        if not quelle:
            raise ValueError(f"Kein {art} angegeben.")
        if ":" in quelle and not os.path.splitdrive(quelle)[0] and not quelle.lower().endswith(".qdat"):
            ok = self.QF.db_object_exists(self.q.PathName(quelle)).value
            self.log(f"{art}: Datenbankpfad {quelle} (vorhanden={ok})", 1)
            if not ok:
                raise RuntimeError(f"Datenbankobjekt nicht gefunden: {quelle}")
            return quelle
        datei = Path(quelle)
        if not datei.is_file():
            raise FileNotFoundError(f"{art}-Datei nicht gefunden: {datei}")
        ziel = f"doc:{art}/{datei.stem}"
        self.QF.db_object_import(self.q.SrcTargetPath(str(datei), ziel))
        ok = self.QF.db_object_exists(self.q.PathName(ziel)).value
        self.log(f"{art} importiert: {datei.name} -> {ziel} (vorhanden={ok})", 1)
        return ziel

    def antrieb_anlegen(self):
        pr = self.cfg["presse"]
        kraft = float(pr["nennkraft_MN"]) * 1e6
        v = float(pr["v_mm_s"])
        name = f"Hydraulikpresse_{pr['nennkraft_MN']:g}MN_{v:g}mm_s".replace(".", ",")
        datei = self.aus / f"{name}.qdat-bin"
        datei.write_bytes(antrieb_bytes(kraft, v))
        ziel = f"doc:drive/{name}"
        self.QF.db_object_import(self.q.SrcTargetPath(str(datei), ziel))
        pruef = []
        for feld in ["hyd_speed", "hyd_max_load"]:
            try:
                pv = self.QF.db_property_get(self.q.DbProperty(ziel, feld))
                pruef.append(f"{feld} = {pv.value} {pv.units}".strip())
            except Exception as ex:
                pruef.append(f"{feld} nicht lesbar ({ex})")
        self.log(f"Antrieb Obersattel: {ziel}  [" + "; ".join(pruef) + "]", 1)
        return ziel

    # ------------------------------------------------------------------ Geometrie Stich 1
    def werkstueck_laden(self):
        wp = self.cfg["werkstueck"]
        g = wp["geometrie"]
        q = self.q
        if g.get("quelle") == "datei":
            datei = Path(g["datei"])
            if not datei.is_file():
                raise FileNotFoundError(f"Werkstückgeometrie nicht gefunden: {datei}")
            self.QF.geometry_load_single_object(q.FileObject(str(datei), self.WP, 1))
            self.log(f"Werkstückgeometrie geladen: {datei.name}", 1)
            return None
        form, d, L, r = g["form"], float(g["d"]), float(g["L"]), float(g.get("r") or 0.0)
        methode = g.get("methode", "auto")
        if methode in ("auto", "grundkoerper") and r <= 0 and form in ("rund", "vier"):
            try:
                if form == "rund":
                    self.QF.geometry_create_tube(q.TubeObjectParams(object_type=self.WP, object_id=1, r=d / 2.0,
                                                                    inner_r=0.0, h=L))
                else:
                    self.QF.geometry_create_brick(q.BrickObjectParams(object_type=self.WP, object_id=1, size_x=d,
                                                                      size_y=d, size_z=L))
                self.log(f"Werkstück aus dem Rohblock erzeugt (Grundkörper): {form} {d:g} x {L:g} mm", 1)
                return "z"
            except Exception as ex:
                if methode == "grundkoerper":
                    raise
                self.log(f"(Grundkörper nicht erzeugt: {ex}; versuche Extrusion)", 1)
        dxf = self.aus / f"{self.name}_Rohblock_Querschnitt.dxf"
        dxf.write_text(querschnitt_dxf(form, d, r), encoding="ascii")
        self.QF.geometry_load_extruded_object(q.ExtrudedObject(file=str(dxf), object_type=self.WP, object_id=1,
                                                               depth=L))
        self.log(f"Werkstück aus dem Rohblock erzeugt (Extrusion {dxf.name}): {form} {d:g} x {L:g} mm"
                 + (f", r = {r:g} mm" if r > 0 else ""), 1)
        return "z"

    def sattel_platzieren(self, oid, oben, z_flaeche):
        """Sattel ausrichten (Breite in x), Arbeitsfläche zum Werkstück, Fläche auf z_flaeche, mittig x = y = 0."""
        b = self.box(self.TOOL, oid)
        self.log(f"Sattel {oid} geladen:  {b}", 1)
        if b.ext[0] > b.ext[1]:
            self.drehen(self.TOOL, oid, b.mitte, (0, 0, 1), 90.0)
            b = self.box(self.TOOL, oid)
        flaeche_oben = self.cfg["werkzeug"].get("arbeitsflaeche", "oben") == "oben"
        if oben == flaeche_oben:
            self.drehen(self.TOOL, oid, b.mitte, (1, 0, 0), 180.0)
            b = self.box(self.TOOL, oid)
        cx, cy, _ = b.mitte
        z_ziel = z_flaeche - b.z0 if oben else z_flaeche - b.z1
        self.verschieben(self.TOOL, oid, -cx, -cy, z_ziel)
        b = self.box(self.TOOL, oid)
        self.log(f"Sattel {oid} ({'oben' if oben else 'unten'}) platziert: {b}", 1)
        return b

    def arbeitsflaeche(self, oid, b, flaeche_oben=True):
        """x-Bereich der ebenen Arbeitsfläche eines Sattels (ohne Kantenradien) aus den Netzknoten der Fläche
        (Untersattel: Fläche oben bei b.z1)."""
        try:
            m = self.QF.mesh_nodes_get(self.q.MeshObjectId(self.TOOL, oid, 0))
            z_ref = b.z1 if flaeche_oben else b.z0
            xs = [x for x, z in zip(m.x, m.z) if abs(z - z_ref) < 0.3]
            if len(xs) >= 3 and max(xs) - min(xs) > 10.0:
                return min(xs), max(xs)
        except Exception as ex:
            self.log(f"(Arbeitsfläche des Sattels nicht bestimmbar: {ex})", 1)
        return b.x0, b.x1

    def werkstueck_ausrichten(self, achse_hint=None):
        """Werkstückachse nach x drehen (längste Ausdehnung bzw. Vorgabe), auf z = 0 legen, y mittig."""
        b = self.box(self.WP, 1)
        self.log(f"Werkstück geladen:  {b}", 1)
        ex, ey, ez = b.ext
        if achse_hint == "z" or (achse_hint is None and ez > max(ex, ey) * 1.05):
            self.drehen(self.WP, 1, b.mitte, (0, 1, 0), 90.0)
        elif achse_hint == "y" or (achse_hint is None and ey > max(ex, ez) * 1.05):
            self.drehen(self.WP, 1, b.mitte, (0, 0, 1), 90.0)
        b = self.box(self.WP, 1)
        self.verschieben(self.WP, 1, 0.0, -b.mitte[1], -b.z0)
        return self.box(self.WP, 1)

    # ------------------------------------------------------------------ Lage des ersten Bisses
    def ziel_x0(self, b, unten, stich):
        """x-Lage des freien Endes (x0) für den ersten Biss eines Stichs: das Ende, an dem die Bissfolge beginnt,
        ragt eine Bisslänge unter den Sattel."""
        fx0, fx1 = self.flach
        s = min(float(stich["bisslaenge"]), fx1 - fx0)
        if self.cfg.get("bissfolge", "manipulator") == "manipulator":
            return fx0 + s - b.ext[0]                # gegriffenes Ende (+x) bei x = Flächenkante (−x) + s
        return fx1 - s                               # freies Ende (−x) bei x = Flächenkante (+x) − s

    def fenster_unter_sattel(self, b, unten, x0_neu):
        """x-Bereich (im aktuellen Koordinatensystem des Werkstücks), der nach dem Verschieben unter dem Sattel liegt."""
        dx = x0_neu - b.x0
        xa, xb = max(b.x0, unten.x0 - dx), min(b.x1, unten.x1 - dx)
        if xb - xa < 5.0:
            xa, xb = b.x0, min(b.x1, b.x0 + 50.0)
        return xa, xb

    def achse_und_manipulator(self, b, cy, cz):
        q, m = self.q, self.cfg["manipulator"]
        self.QF.axis_set(q.ObjectAxisParams(self.WP, 1, 1, b.x0, cy, cz, b.x0 + 1000.0, cy, cz))
        bc = self.QF.bound_cond_create(q.BoundCondParams(id=-1, type=q.BCond.Manipulator, object_type=self.WP,
                                                         object_id=1))
        bid = bc.id
        lg = float(m["grifflaenge"])
        hx = (lg + 20.0) / 2
        hy = max(b.y1 - cy, cy - b.y0) + 20.0
        hz = max(b.z1 - cz, cz - b.z0) + 20.0
        self.QF.bound_cond_set_brick(q.ShapeBrick(bid, hx, hy, hz, hx, hy, hz))
        self.verschieben(q.ObjectType.BoundaryCondition, bid, b.x1 + 20.0 - hx, cy, cz)
        BC = q.ObjectType.BoundaryCondition
        self.wert(BC, bid, "params/give_back", "true" if m.get("rueckstellung", True) else "false")
        self.wert(BC, bid, "params/axial_resistance_factor", m.get("widerstand_axial", 1.5))
        self.wert(BC, bid, "params/crosswise_resistance_factor", m.get("widerstand_quer", 1.0))
        self.log(f"Achse x bei y={cy:.1f} z={cz:.1f}; Manipulator {bid}: x {b.x1 - lg:.1f} .. {b.x1 + 20:.1f} "
                 f"(Grifflänge {lg:g} mm)", 1)
        return bid

    def stoppbedingung(self, mass):
        q = self.q
        sc = self.QF.stop_cond_create(q.StopCondParams(id=-1, type=q.StopCondType.Distance))
        sid = sc.id
        SC = q.ObjectType.StopCondition
        self.objekt(SC, sid, "data/object_1", self.TOOL, 1)
        self.objekt(SC, sid, "data/object_2", self.TOOL, 2)
        self.wert(SC, sid, "data/fin_dist", f"{mass:g}")
        self.wert(SC, sid, "for_blow", "false")
        self.log(f"Stoppbedingung {sid}: Abstand der Sättel {mass:g} mm", 1)

    def operation_einstellen(self, op_id, k, stich):
        wp, wz = self.cfg["werkstueck"], self.cfg["werkzeug"]
        self.objekt(self.OP, op_id, "tool_put_in_contact_ref_object", self.WP, 1)
        mikro = bool(wp.get("mikrostruktur"))
        self.wert(self.OP, op_id, "with_microstructure", "true" if mikro else "false")
        if k == 0:
            self.wert(self.WP, 1, "params/material_id", self.wp_werkstoff)
            if mikro:
                self.wert(self.WP, 1, "params/grain_size", f"{float(wp.get('korngroesse_um', 100)):g}")
            self.wert(self.WP, 1, "params/temper", f"{float(wp['temperatur']):g}")
        else:
            self.wert(self.WP, 1, "params/grain_inherited", "true")
            if stich.get("waermen") is not None:
                self.wert(self.WP, 1, "params/temper_inherited", "false")
                self.wert(self.WP, 1, "params/temper", f"{float(stich['waermen']):g}")
            else:
                self.wert(self.WP, 1, "params/temper_inherited", "true")
        vererben = bool(wz.get("temperatur_vererben"))
        for tid, antrieb in [(1, ANTRIEB_FEST), (2, self.antrieb)]:
            self.wert(self.TOOL, tid, "drive_id", antrieb)
            self.wert(self.TOOL, tid, "params/lubricant_id", wz["schmierstoff"])
            self.wert(self.TOOL, tid, "params/material_id", self.wz_werkstoff)
            if k > 0:
                self.wert(self.TOOL, tid, "params/temper_inherited", "true" if vererben else "false")
            if k == 0 or not vererben:
                self.wert(self.TOOL, tid, "params/temper", f"{float(wz['temperatur']):g}")

    def huebe_setzen(self, blows, manip_id):
        q = self.q
        self.QF.blow_count_set(q.Count(len(blows)))
        aktiv = bool(self.cfg["manipulator"].get("aktiv_im_hub", False))
        for i, (luft, werkzeug, vorschub, drehung) in enumerate(blows, start=1):
            for p, w in [(q.BlowParam.CoolingInAir, luft), (q.BlowParam.CoolingInTool, werkzeug),
                         (q.BlowParam.Feed, vorschub), (q.BlowParam.Rotation, drehung)]:
                self.QF.blow_parameter_set(q.BlowParameter(blow=i, param=p, value=float(w)))
            self.QF.blow_manipulator_set_active(q.BlowManipulatorIsActive(blow=i, manipulator=manip_id,
                                                                          is_active=aktiv))
        vorsch = [b[2] for b in blows if abs(b[2]) > 1e-9]
        self.log(f"{len(blows)} Hübe übernommen: vor Hub 1 {blows[0][0]:g} s an Luft, je Hub {blows[0][1]:g} s auf "
                 f"dem Sattel, {len(vorsch)} Vorschübe" + (f" à {vorsch[0]:+.1f} mm" if vorsch else ""), 1)

    # ------------------------------------------------------------------ Stich anlegen
    def stich_anlegen(self, k, stich, vorgaenger, b_vorher):
        q = self.q
        name = stich["name"]
        self.log("")
        self.log(f"=== {name}: Maß {stich['mass']:g} mm, Länge laut Stichplan {stich['laenge']:.0f} mm, "
                 f"{len(stich['blows'])} Hübe ===")
        self.melde(zustand="aufbau", stich=k + 1, hub=0, huebe=len(stich["blows"]),
                   meldung=f"{name} wird aufgebaut")
        gap = float(self.cfg["werkzeug"].get("spalt", 10.0))
        if vorgaenger is None:
            op = self.QF.operation_create(q.OperationParams(
                id=-1, name=name, creation_mode=q.OperationCreationMode.CreateAsNewProcess,
                process_name=self.cfg.get("prozessname") or self.name, operation_type=q.OperationType.deformation,
                problem_type=q.ProblemType.General3D))
            op_id = op.id
            sattel = self.cfg["werkzeug"]["geometrie"]
            if not Path(sattel).is_file():
                raise FileNotFoundError(f"Sattelgeometrie nicht gefunden: {sattel}")
            self.QF.geometry_load_single_object(q.FileObject(str(sattel), self.TOOL, 1))
            self.QF.geometry_load_single_object(q.FileObject(str(sattel), self.TOOL, 2))
            achse_hint = self.werkstueck_laden()
            self.wert(self.TOOL, 1, "dsc", "Untersattel")
            self.wert(self.TOOL, 2, "dsc", "Obersattel")
            self.wert(self.WP, 1, "dsc", self.cfg["werkstueck"].get("bezeichnung", "Rohblock"))
            self.wp_werkstoff = self.db_objekt(self.cfg["werkstueck"]["werkstoff"], "w-material")
            self.wz_werkstoff = self.db_objekt(self.cfg["werkzeug"]["werkstoff"], "t-material",
                                               "std:t-material/H11 HRC51")
            self.antrieb = self.antrieb_anlegen()
            unten = self.sattel_platzieren(1, oben=False, z_flaeche=0.0)
            self.sattelbreite = unten.ext[0]
            self.flach = self.arbeitsflaeche(1, unten)
            w = self.flach[1] - self.flach[0]
            self.log(f"Ebene Arbeitsfläche des Untersattels: x {self.flach[0]:.1f} .. {self.flach[1]:.1f} "
                     f"(Breite {w:.1f} mm, Sattel {self.sattelbreite:.1f} mm)", 1)
            s_max = max(float(x["bisslaenge"]) for x in self.cfg["stiche"])
            if s_max > w + 1.0:
                self.log(f"WARNUNG: Bisslänge bis {s_max:g} mm, ebene Arbeitsfläche nur {w:.0f} mm "
                         f"(Kantenradien). Zwischen den Bissen bleiben Bereiche ungeschmiedet; im Stichplan die "
                         f"Sattelbreite auf {w:.0f} mm setzen.", 1)
            b = self.werkstueck_ausrichten(achse_hint)
            x0 = self.ziel_x0(b, unten, stich)
            self.verschieben(self.WP, 1, x0 - b.x0, 0.0, 0.0)
            b = self.box(self.WP, 1)
            self.log(f"Werkstück platziert: {b}", 1)
            xa, xb = max(b.x0, unten.x0), min(b.x1, unten.x1)
            h_unter = b.z1
            try:
                f = fenster(self.knoten(), xa, xb)
                if f:
                    h_unter = f[3]
            except Exception:
                pass
            self.sattel_platzieren(2, oben=True, z_flaeche=h_unter + gap)
            cy, cz = b.mitte[1], b.ext[2] / 2.0
        else:
            op = self.QF.operation_create(q.OperationParams(
                id=-1, name=name, parent=vorgaenger, creation_mode=q.OperationCreationMode.InsertAfterParent,
                operation_type=q.OperationType.deformation, problem_type=q.ProblemType.General3D))
            op_id = op.id
            for typ, oid in [(self.WP, 1), (self.TOOL, 1), (self.TOOL, 2)]:
                self.QF.object_inherit(q.ObjectId(typ, oid))
            try:
                b = self.box(self.WP, 1)
                self.log(f"geerbtes Werkstück: {b}", 1)
            except Exception as ex:
                b = b_vorher
                self.log(f"(Bounding-Box des geerbten Werkstücks nicht lesbar: {ex}; nehme Endzustand Vorstich)", 1)
            unten = self.box(self.TOOL, 1)
            x0 = self.ziel_x0(b, unten, stich)
            xa, xb = self.fenster_unter_sattel(b, unten, x0)
            sel = None
            try:
                sel = fenster(self.knoten(), xa, xb)
            except Exception as ex:
                self.log(f"(Knoten des geerbten Werkstücks nicht lesbar: {ex})", 1)
            if sel is None:
                sel = (b.y0, b.y1, b.z0, b.z1)
                self.log("(Querschnitt unter dem Sattel nicht bestimmbar, nehme Gesamtbox)", 1)
            dy, dz = -(sel[0] + sel[1]) / 2, -sel[2]
            self.verschieben(self.WP, 1, x0 - b.x0, dy, dz)
            b = Box(q.BoundBox(x0, b.y0 + dy, b.z0 + dz, x0 + b.ext[0], b.y1 + dy, b.z1 + dz))
            h_unter = sel[3] - sel[2]
            self.log(f"Werkstück für den ersten Biss positioniert: Querschnitt unter dem Sattel "
                     f"{sel[1] - sel[0]:.1f} x {h_unter:.1f} mm auf z = 0 | Gesamt: {b}", 1)
            oben = self.box(self.TOOL, 2)
            if abs(oben.z0 - (h_unter + gap)) > 1.0:
                self.verschieben(self.TOOL, 2, 0.0, 0.0, h_unter + gap - oben.z0)
                self.log(f"Obersattel auf z = {h_unter + gap:.1f} mm gesetzt ({gap:g} mm Luft)", 1)
            cy, cz = 0.0, h_unter / 2.0
        w = self.flach[1] - self.flach[0]
        sb = float(self.cfg.get("sattelbreite_stichplan", w))
        if k == 0 and abs(sb - w) > 1.0 and abs(sb - self.sattelbreite) > 1.0:
            self.log(f"WARNUNG: Sattelbreite im Stichplan {sb:g} mm, Sattelgeometrie {self.sattelbreite:.1f} mm "
                     f"(ebene Fläche {w:.1f} mm). Bisse und Vorschübe passen nicht zum Sattel.", 1)
        manip = self.achse_und_manipulator(b, cy, cz)
        self.stoppbedingung(float(stich["mass"]))
        self.operation_einstellen(op_id, k, stich)
        self.huebe_setzen(stich["blows"], manip)
        if k > 0 and stich.get("waermen") is not None:
            self.log(f"Wiedererwärmung: Werkstücktemperatur nicht vererbt, homogen {float(stich['waermen']):g} °C", 1)
        self.log(f"Operation {op_id} angelegt.", 1)
        return op_id, b

    # ------------------------------------------------------------------ Rechnen und Auswerten
    def stich_rechnen(self, op_id, k, n_huebe):
        q = self.q
        self.QF.operation_set_current(q.ItemId(op_id))
        t0 = time.time()
        self.QF.async_start_simulation(q.SimulationParams(calculation_mode=q.CalculationMode.Operation))
        max_s = float(self.cfg.get("max_rechenzeit_h", 8.0)) * 3600
        letzte, letztes, aussetzer = 0.0, None, 0
        while True:
            if self.stop_pfad.exists():
                self.log("Abbruch angefordert: Berechnung wird angehalten.", 1)
                try:
                    self.QF.async_stop_simulation()
                except Exception:
                    pass
                raise Abbruch("Abbruch angefordert")
            try:
                e = self.QF.async_wait(q.AsyncWaitingParams(timeout=10))
                aussetzer = 0
            except Exception as ex:
                s = self.QF.async_state()
                aussetzer += 1
                if s.working and aussetzer <= 10:
                    time.sleep(2)
                    continue
                if not s.working:
                    letztes = None
                    break
                raise RuntimeError(f"async_wait: {ex}")
            typ = self.ev(e.type)
            jetzt = time.time() - t0
            self.melde(zustand="rechnet", stich=k + 1, hub=int(getattr(e, "blow", 0) or 0), huebe=n_huebe,
                       prozesszeit=float(getattr(e, "process_time", 0.0) or 0.0),
                       meldung=f"Stich {k + 1}: Hub {getattr(e, 'blow', '?')} von {n_huebe}, {jetzt / 60:.1f} min")
            if typ == "NextCalculationUnit" or jetzt - letzte > 120 or typ in ENDE:
                self.log(f"{jetzt / 60:6.1f} min  {typ:<20} Hub {e.blow}  Record {e.record:g}  "
                         f"Prozesszeit {e.process_time:.2f} s", 1)
                letzte = jetzt
            letztes = e
            if typ in ENDE or (typ == "Idle" and not e.working):
                break
            if jetzt > max_s:
                self.QF.async_stop_simulation()
                raise RuntimeError(f"Rechenzeit-Grenze {max_s / 3600:g} h erreicht")
        typ = self.ev(letztes.type) if letztes is not None else "unbekannt"
        st = self.QF.state_operation(q.OptionalItemId(op_id))
        if typ in ("Error", "HaspKeyNotFound", "Canceled") or not st.simulation_completed:
            raise RuntimeError(f"Simulation nicht abgeschlossen (Ereignis {typ}, abgeschlossen="
                               f"{st.simulation_completed}): {getattr(letztes, 'diagnostic_msg', '')}")
        return time.time() - t0

    def presskraft_max(self, n_huebe, tool_id=2):
        """Größte Kraft am Obersattel je Hub (Records zählen je Hub) und insgesamt."""
        q = self.q
        f_max, einheit, je_hub = None, "", []
        for hub in range(1, int(n_huebe) + 1):
            try:
                self.QF.blow_set_current(q.ItemId(hub))
                rec = int(self.QF.record_get_last().record)
            except Exception:
                if hub == 1:
                    raise
                break
            f_hub = None
            for r in range(1, rec + 1):
                z = self.QF.invoke(0, "state_tool", q.ToolStateId(record=r, id=tool_id, type=self.TOOL),
                                   _WerkzeugZustand)
                if z is not None and z.kraft is not None and (f_hub is None or abs(z.kraft) > f_hub):
                    f_hub = abs(z.kraft)
                    einheit = z.einheiten or einheit
            je_hub.append(None if f_hub is None else round(f_hub, 3))
            if f_hub is not None and (f_max is None or f_hub > f_max):
                f_max = f_hub
        for teil in (einheit or "").split(";"):
            if teil.startswith("load:"):
                einheit = teil[5:]
        return f_max, einheit, je_hub

    def stich_auswerten(self, op_id, k, stich, dauer):
        q = self.q
        rec = self.QF.record_get_last().record
        self.QF.record_set(q.Record(rec))
        b = self.box(self.WP, 1)
        f_max, einheit, f_hub = None, "", []
        try:
            f_max, einheit, f_hub = self.presskraft_max(len(stich["blows"]))
        except Exception as ex:
            self.log(f"(Presskraft nicht auslesbar: {ex})", 1)
        finally:
            try:                                     # Endzustand wieder einstellen (letzter Hub, letzter Record)
                self.QF.blow_set_current(q.ItemId(len(stich["blows"])))
                self.QF.record_set(q.Record(self.QF.record_get_last().record))
            except Exception:
                pass
        prof, mitte = [], None
        try:
            pk = self.knoten()
            prof = profil(pk)
            mitte = fenster(pk, b.x0 + 0.4 * b.ext[0], b.x0 + 0.6 * b.ext[0])
        except Exception as ex:
            self.log(f"(Knotenprofil nicht auslesbar: {ex})", 1)
        mass = float(stich["mass"])
        l_auf_mass = sum(20.0 for _x, _bb, hh in prof if hh <= mass + 3.0) if prof else None
        erg = {"stich": k + 1, "operation": op_id, "name": stich["name"], "mass_soll": mass,
               "laenge_stichplan": round(float(stich["laenge"]), 1), "laenge_ist": round(b.ext[0], 1),
               "querschnitt_mitte": None if mitte is None else [round(mitte[1] - mitte[0], 1),
                                                                round(mitte[3] - mitte[2], 1)],
               "laenge_auf_mass": l_auf_mass, "huelle": [round(b.ext[1], 1), round(b.ext[2], 1)],
               "f_max": None if f_max is None else round(f_max, 3), "f_einheit": einheit,
               "f_stichplan_MN": stich.get("f_plan_MN"), "f_je_hub": f_hub, "records": rec, "rechenzeit_min": round(dauer / 60, 1),
               "huebe": len(stich["blows"]), "profil_x_b_h": prof}
        self.ergebnisse.append(erg)
        qm = erg["querschnitt_mitte"]
        self.log(f"Ergebnis: L = {b.ext[0]:.0f} mm (Stichplan {float(stich['laenge']):.0f}), Querschnitt Mitte "
                 f"{qm[0] if qm else '?'} x {qm[1] if qm else '?'} mm, Hülle {b.ext[1]:.1f} x {b.ext[2]:.1f} mm, "
                 f"F_max = {erg['f_max']} {einheit}, {dauer / 60:.1f} min", 1)
        if f_hub:
            self.log("F_max je Hub [" + einheit + "]: " + ", ".join("–" if f is None else f"{f:.2f}" for f in f_hub), 1)
        return b

    # ------------------------------------------------------------------ Hauptablauf
    def ausfuehren(self):
        self.aus.mkdir(parents=True, exist_ok=True)
        self.log_pfad.parent.mkdir(parents=True, exist_ok=True)
        self.status_pfad.parent.mkdir(parents=True, exist_ok=True)
        try:
            if self.stop_pfad.exists():
                self.stop_pfad.unlink()
        except OSError:
            pass
        self.logdatei = open(self.log_pfad, "w", encoding="utf-8")
        cfg = self.cfg
        self.log(f"Freiformschmieden-Rechner {cfg.get('programmversion', '')}: QForm-Simulation "
                 f"{'(nur Aufbau Stich 1)' if self.nur_aufbau else '(Aufbau und Berechnung)'}")
        self.log(f"Werkstoff {cfg['werkstueck'].get('bezeichnung', '')}, {len(cfg['stiche'])} Stiche, Bissfolge "
                 f"{'vom Manipulator weg' if cfg.get('bissfolge', 'manipulator') == 'manipulator' else 'zum Manipulator hin'}"
                 f", Presse {cfg['presse']['nennkraft_MN']:g} MN / {cfg['presse']['v_mm_s']:g} mm/s, "
                 f"Schmierstoff {cfg['werkzeug']['schmierstoff']}")
        self.melde(zustand="startet", meldung="QForm wird gestartet")
        fehler, abgebrochen = None, False
        try:
            self.qform_starten()
            vorgaenger, b_vorher = None, None
            for k, stich in enumerate(cfg["stiche"]):
                self.stop_pruefen()
                op_id, b_vorher = self.stich_anlegen(k, stich, vorgaenger, b_vorher)
                if vorgaenger is None:
                    self.QF.project_save_as(self.q.FileName(str(self.projekt)))
                    self.projekt_angelegt = True
                    self.log(f"Projekt gespeichert: {self.projekt}", 1)
                else:
                    self.QF.project_save()
                if self.nur_aufbau:
                    self.log("Nur Aufbau: Stich 1 angelegt, keine Berechnung. QForm bleibt zur Kontrolle offen.")
                    break
                self.log("Berechnung läuft ...", 1)
                dauer = self.stich_rechnen(op_id, k, len(stich["blows"]))
                b_vorher = self.stich_auswerten(op_id, k, stich, dauer)
                self.QF.project_save()
                self.melde(meldung=f"Stich {k + 1} fertig")
                vorgaenger = op_id
        except Abbruch as ex:
            abgebrochen = True
            fehler = str(ex)
            self.log(f"ABBRUCH: {ex}")
        except Exception as ex:
            fehler = str(ex) or ex.__class__.__name__
            self.log(f"FEHLER: {fehler}")
            for z in traceback.format_exc().rstrip().splitlines():
                self.log(z, 1)
        finally:
            if self.QF is not None and self.projekt_angelegt:
                try:
                    self.QF.project_save()
                except Exception:
                    pass
            with open(self.erg_pfad, "w", encoding="utf-8") as fh:
                json.dump({"zeit": self.zeit, "projekt": str(self.projekt), "fehler": fehler,
                           "nur_aufbau": self.nur_aufbau, "stiche": self.ergebnisse}, fh, ensure_ascii=False, indent=1)
            self.log("")
            self.log("=" * 80)
            self.log("ZUSAMMENFASSUNG")
            for e in self.ergebnisse:
                qm = e["querschnitt_mitte"] or ["?", "?"]
                self.log(f"Stich {e['stich']}: {e['huebe']} Hübe, L {e['laenge_ist']} mm (Stichplan "
                         f"{e['laenge_stichplan']:g}), Mitte {qm[0]} x {qm[1]} mm (Maß {e['mass_soll']:g}), "
                         f"F_max {e['f_max']} {e['f_einheit']} (Stichplan {e['f_stichplan_MN']} MN), "
                         f"{e['rechenzeit_min']} min", 1)
            self.log(f"Ergebnis: {'ABGEBROCHEN' if abgebrochen else ('FEHLER - ' + fehler if fehler else 'OK')}")
            self.log(f"Projekt: {self.projekt if self.projekt_angelegt else '(nicht angelegt)'}")
            self.log("=" * 80)
            if self.QF is not None:
                try:
                    if not self.qform_bereit:
                        self.QF.qform_close()          # nichts zu kontrollieren (z. B. keine Lizenz)
                        self.log("QForm geschlossen.")
                    elif self.nur_aufbau or fehler or not cfg.get("qform_am_ende_schliessen", True):
                        self.QF.qform_detach()
                        self.log("QForm bleibt geöffnet.")
                    else:
                        self.QF.qform_close()
                        self.log("QForm geschlossen.")
                except Exception:
                    pass
            self.melde(zustand="abgebrochen" if abgebrochen else ("fehler" if fehler else "fertig"),
                       meldung=fehler or ("Aufbau fertig" if self.nur_aufbau else "Berechnung fertig"))
            try:
                if self.stop_pfad.exists():
                    self.stop_pfad.unlink()
            except OSError:
                pass
            self.logdatei.close()
        return 1 if fehler else 0


# ======================================================================== Geometrie-Hilfen
class Box:
    def __init__(self, b):
        self.x0, self.y0, self.z0 = b.min_x, b.min_y, b.min_z
        self.x1, self.y1, self.z1 = b.max_x, b.max_y, b.max_z

    @property
    def ext(self):
        return (self.x1 - self.x0, self.y1 - self.y0, self.z1 - self.z0)

    @property
    def mitte(self):
        return ((self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2, (self.z0 + self.z1) / 2)

    def gueltig(self):
        return all(e > 1.0 for e in self.ext)

    def __str__(self):
        return (f"x {self.x0:8.1f} .. {self.x1:8.1f} | y {self.y0:7.1f} .. {self.y1:7.1f} | "
                f"z {self.z0:7.1f} .. {self.z1:7.1f}  (L x B x H = {self.ext[0]:.1f} x {self.ext[1]:.1f}"
                f" x {self.ext[2]:.1f} mm)")


def fenster(pkte, xa, xb):
    """Querschnittsgrenzen (y0, y1, z0, z1) aller Knoten mit xa <= x <= xb, None bei zu wenigen Knoten."""
    sel = [(y, z) for x, y, z in pkte if xa <= x <= xb]
    if len(sel) < 8:
        return None
    ys, zs = [p[0] for p in sel], [p[1] for p in sel]
    return min(ys), max(ys), min(zs), max(zs)


def profil(pkte, schritt=20.0):
    """Längsprofil: Liste (x ab freiem Ende, Breite, Höhe) in Scheiben der Dicke 'schritt'."""
    x0, x1 = min(p[0] for p in pkte), max(p[0] for p in pkte)
    erg, xa = [], x0
    while xa < x1:
        f = fenster(pkte, xa, xa + schritt)
        if f:
            erg.append((round(xa + schritt / 2 - x0, 1), round(f[1] - f[0], 1), round(f[3] - f[2], 1)))
        xa += schritt
    return erg


class _WerkzeugZustand:
    """Ersatz für QFormAPI.ToolState: dort überdeckt das Datenfeld 'load' die Methode load()."""

    def load(self, B):
        self.record = B.prop_get("record")
        self.kraft = B.prop_get("load")
        self.einheiten = B.prop_get("units")


# ======================================================================== Aufruf
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    nur_aufbau = "--nur-aufbau" in argv

    def opt(name):
        if name in argv:
            i = argv.index(name)
            if i + 1 < len(argv):
                return argv[i + 1]
        return None
    pos = [a for i, a in enumerate(argv) if not a.startswith("--") and (i == 0 or argv[i - 1] not in
                                                                         ("--log", "--status", "--stop"))]
    if not pos:
        print(__doc__)
        return 2
    try:
        cfg = lade_konfiguration(pos[0])
        lauf = Lauf(cfg, nur_aufbau=nur_aufbau, log_pfad=opt("--log"), status_pfad=opt("--status"),
                    stop_pfad=opt("--stop"))
    except Exception as ex:                       # Fehler vor dem Start: im Status melden, damit die Oberfläche ihn zeigt
        if opt("--status"):
            try:
                with open(opt("--status"), "w", encoding="utf-8") as fh:
                    json.dump(dict(zustand="fehler", meldung=f"Konfiguration nicht lesbar: {ex}", ergebnisse=[]), fh,
                              ensure_ascii=False)
            except OSError:
                pass
        if sys.stdout is not None:
            print(f"FEHLER: {ex}")
        return 1
    return lauf.ausfuehren()


if __name__ == "__main__":
    sys.exit(main())
