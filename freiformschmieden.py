# -*- coding: utf-8 -*-
"""
Freiformschmieden Aluminium – Vorberechnung
===========================================
Desktop-Programm (Tkinter) für
  * Endlänge nach dem Schmieden (Volumenkonstanz) für Rund-, Vierkant- und Achtkantquerschnitte
  * Aufheizzeit im Ofen (instationäre Wärmeleitung, Konvektion + Strahlung)

Nur Python-Standardbibliothek; als Einzeldatei-.exe mit PyInstaller paketierbar.
"""
import csv
import datetime
import json
import math
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import time
import tkinter.font as tkfont

import stichplan as sp
import qform_export
import qform_db
import qform_runner
import qform_sim

APP_TITLE = "Freiformschmieden Aluminium – Vorberechnung"
VERSION = "1.6"

# ---------------------------------------------------------------- Daten
# kf: Sellars-Tegart (alpha [1/MPa], n, Q [kJ/mol]) und Anker k_f(T_ref, 1 s⁻¹) [MPa] -> lnA.
# Grobe Richtwerte für den stationären Bereich (Unsicherheit etwa ±30 %), eigene Daten bevorzugen.
MATS = [
    dict(id="6082", name="EN AW-6082", rho=2700, cp=900, k=180, kf=(0.045, 3.6, 160, 450, 58)),
    dict(id="6061", name="EN AW-6061", rho=2700, cp=896, k=167, kf=(0.045, 3.6, 160, 450, 55)),
    dict(id="7075", name="EN AW-7075", rho=2810, cp=960, k=130, kf=(0.025, 4.0, 150, 450, 80)),
    dict(id="2014", name="EN AW-2014", rho=2800, cp=880, k=155, kf=(0.030, 4.0, 155, 450, 75)),
    dict(id="2618", name="EN AW-2618A", rho=2760, cp=875, k=150, kf=(0.030, 4.0, 160, 450, 80)),
    dict(id="5083", name="EN AW-5083", rho=2660, cp=900, k=120, kf=(0.030, 3.8, 170, 450, 90)),
    dict(id="1050", name="EN AW-1050A", rho=2700, cp=900, k=229, kf=(0.050, 4.5, 156, 450, 28)),
    dict(id="az31", name="AZ31B (Mg)", rho=1770, cp=1000, k=96, kf=(0.020, 5.0, 140, 400, 55)),
    dict(id="az80", name="AZ80A (Mg)", rho=1800, cp=1000, k=76, kf=(0.020, 4.5, 150, 400, 75)),
]


def fe(x, d=0):
    """Zahl für Eingabefelder: Dezimalkomma, ohne Tausendertrennzeichen."""
    return f"{x:.{d}f}".replace(".", ",")


def fg(x):
    return f"{x:.6g}".replace(".", ",")


def kf_fields(m):
    """Eingabewerte der Fließspannung (Sellars-Tegart und angepasster Hensel-Spittel) für einen Werkstoff."""
    a, n, Q, Tr, sr = m["kf"]
    lnA = sp.st_lnA_from_anchor(a, n, Q, Tr, sr)
    hs = sp.hs_from_st(dict(alpha=a, n=n, Q=Q, lnA=lnA), Tr - 100, Tr + 100)
    out = dict(kfAlpha=fg(a), kfN=fg(n), kfQ=fg(Q), kfLnA=fe(lnA, 3))
    for key in sp_HS_KEYS:
        out["hs" + key] = fg(hs[key])
    return out


sp_HS_KEYS = ("A", "m1", "m2", "m3", "m4", "m5", "m7", "m8", "m9")
SHAPE_BY_NAME = {"Rund": "rund", "Vierkant": "vier", "Achtkant": "acht"}
CUSTOM = "Benutzerdefiniert"
FURN = [
    dict(name="Luftumwälzofen", hc=35,
         hint="Erzwungene Konvektion, Richtwert 20 bis 60 W/m²K."),
    dict(name="Hochkonvektions- oder Düsenofen", hc=100,
         hint="Angeblasene Oberflächen, 60 bis 150 W/m²K."),
    dict(name="Kammerofen ohne Umwälzung", hc=8,
         hint="Freie Konvektion 5 bis 10 W/m²K, sonst Strahlung."),
]
SHAPES = {
    "rund": dict(name="Rund", dim="Durchmesser D", pre="Ø"),
    "vier": dict(name="Vierkant", dim="Kantenlänge a", pre="□"),
    "acht": dict(name="Achtkant", dim="Schlüsselweite SW", pre="SW"),
}
DEFAULTS = dict(
    matPreset="EN AW-6082", matName="EN AW-6082", rho="2700", cp="900", lam="180",
    blkShape="rund", blkD="200", blkR="0", blkL="400",
    tgtShape="vier", tgtMode="dim", tgtD="120", tgtL="1000", tgtR="0", loss="0",
    heatGeo="blk", T0="20", Tf="500", Tz="490", furnPreset="Luftumwälzofen", hc="35", eps="0,1",
    pT0="480", pTrans="20", pAux="5", pPause="10",
    pV="60", pDwell="0,2", pBdie="250", pSB="0,6", pBite="min", pSBc="150",
    qfScope="stich", qfFormat="xlsx", qfDec="punkt", qfSign="neg", qfReturn="1",
    pTdie="250", pHc="15000", pHrest="3000", pLdie="30", pHair="10", pEps="0,1",
    kfModel="st", pEta="0,75", pMu="0,3",
    pFmax="10", pTmin="400", pTmax="520",
    planView="temp", planParams="1",
    qsQDir="", qsOut="", qsName="", qsWpGeo="rohblock", qsWpFile="", qsWpMat=qform_sim.DEFAULT_WP_MATERIAL,
    qsWpTm="plan", qsWpT="480", qsMicro="0", qsGrain="100", qsToolGeo=qform_sim.DEFAULT_TOOL, qsToolFace="oben",
    qsToolMat="std:t-material/H11 HRC51", qsToolTm="plan", qsToolT="250", qsToolInh="0",
    qsLub=qform_db.LUBRICANTS[0][0], qsLubPath="", qsPressF="10", qsDir="manipulator", qsExtra="0", qsGrip="80",
    qsElong="0,5", qsResAx="1,5", qsResCr="1", qsGive="1", qsGap="10", qsMaxH="8", qsClose="1", qsLast="",
    tab="0",
)
DEFAULT_PASSES = [
    dict(shape="vier", mode="dim", dim="160", len="", nrot="", bk=""),
    dict(shape="vier", mode="dim", dim="130", len="", nrot="", bk=""),
    dict(shape="acht", mode="dim", dim="120", len="", nrot="", bk=""),
    dict(shape="rund", mode="dim", dim="110", len="", nrot="", bk=""),
]

MODEL_TEXT = """Endlänge
Volumenkonstanz: L1 = A0 · L0 · (1 − Verlust) / A1.
Achtkant als regelmäßiges Achteck mit Schlüsselweite SW: A = 2(√2 − 1) · SW².
Kantenradien verringern den Querschnitt beim Vierkant um (4 − π) · r², beim Achtkant um 8(√2 − 1 − π/8) · r².
Alle Maße sind Kaltmaße; die Wärmedehnung ist nicht berücksichtigt.

Aufheizen
Instationäre Wärmeleitung mit konstanten Stoffwerten, gelöst mit einem impliziten Finite-Volumen-Verfahren. Der Block wird als Produkt eindimensionaler Lösungen behandelt (Newman-Ansatz): Rundblock = Zylinder × Platte, Vierkant = Platte × Platte × Platte. Der Achtkant wird als Zylinder mit r = SW/2 gerechnet, der dasselbe Verhältnis von Volumen zu Mantelfläche hat; gegenüber einer 2D-Rechnung des Achtecks liegt die Abweichung bei etwa 1 %.

Randbedingung
Allseitiger Wärmeübergang durch Konvektion (h) und Strahlung gegen die Ofenwand mit Ofentemperatur (ε, σ = 5,67·10⁻⁸ W/m²K⁴). Der Ofen ist vorgeheizt und hält seine Temperatur. Gegenüber einer 2D-rotationssymmetrischen Rechnung mit örtlich aufgelöster Strahlung weicht die Kernaufheizzeit um weniger als 1 % ab.

Nicht enthalten
Auflageflächen auf Rost oder Herd, gegenseitige Abschattung mehrerer Blöcke und der Temperatureinbruch beim Chargieren. Die Rechnung liefert damit eine Untergrenze der Aufheizzeit für einen einzeln liegenden Block.

Stoffwerte
Die Vorlagen sind Richtwerte bei Raumtemperatur. Die Wärmekapazität von Al-Legierungen steigt bis 500 °C um rund 20 %; für die Aufheizrechnung ist ein Mittelwert über den Temperaturbereich sinnvoll.

Stichplan: Geometrie
Stich 0 ist der Rohblock. Je Stich wird entweder das Maß (D, a oder SW) oder die Länge vorgegeben, das andere folgt aus der Volumenkonstanz. Die Stiche sind scharfkantig gerechnet. Umformgrad je Stich φ = ln(A_vorher/A_nachher). Gerechnet wird Recken mit Flachsätteln in der Hubfolge von QForm: Je Biss werden alle Drehlagen geschmiedet (Drehschritt 180°/Drehlagen), dann folgt der Vorschub zum nächsten Biss. Die größte Bisslänge ist wahlweise die Sattelbreite (minimale Bisszahl), ein konstanter Wert für alle Stiche oder (s_B/h) · h_vorher, jeweils höchstens die Sattelbreite. Die Zahl der Bisse ist L_vorher / größte Bisslänge (aufgerundet); alle Bisse haben die größte Bisslänge, der letzte nimmt den Rest auf; ein Rest unter einem Viertel der Bisslänge wird dem letzten vollen Biss zugeschlagen, sofern dieser dann noch unter den Sattel passt. Der Vorschub zwischen den Bissen ist Bisslänge × A_vorher/A_nachher: Der geschmiedete Abschnitt ist um diesen Faktor länger geworden, damit trifft der Sattel beim nächsten Biss genau auf ungeschmiedeten Werkstoff. Der Hubweg je Hub ist die Ausdehnung des vorherigen Querschnitts in den Druckrichtungen des neuen Stichs minus das neue Maß (beim Übergang Vierkant → Achtkant gemittelt über Flächen und Ecken, bei Rund über den mittleren Durchmesser). Die Kontaktbreite je Sattel ist beim Vierkant die Kantenlänge, beim Achtkant die Seitenlänge und beim Rund die Abplattungsbreite; sie kann je Stich überschrieben werden.

Stichplan: Temperatur
Betrachtet wird eine Scheibe in Stabmitte. Sie liegt im mittleren Biss und wird dort in jeder Drehlage einmal von Ober- und Untersattel gefasst; vor jedem dieser Hübe liegt sie während der Nebenzeit auf dem Untersattel (eine Sattelfläche, Kontaktleitwert h_K beim Aufliegen). Während der übrigen Bisse kühlt sie an Luft. Die Zeitachse entspricht der QForm-Hubtabelle: vor dem ersten Hub eines Stichs Transport bzw. Pause an Luft, vor jedem Hub die Nebenzeit auf dem Untersattel, je Hub die Kontaktzeit. Die Wärmeleitung im Querschnitt wird mit einem Zylinder gleichen Verhältnisses von Volumen zu Oberfläche (R = 2A/U) gerechnet; Kern und Oberfläche beziehen sich auf diesen Ersatzzylinder, die Oberfläche ist der Mittelwert über den Umfang.
Kontaktzeit je Hub = Hubweg / Hubgeschwindigkeit + Verweilzeit. Der Wärmeübergang in den Sattel folgt der geschlossenen Lösung für zwei halbunendliche Körper mit Kontaktleitwert h_K (Wärmeeindringkoeffizienten √(λρc) von Werkstück und Sattel, ρc des Sattels 3,6 MJ/(m³K)). Daraus ergeben sich die Temperatur der Kontaktfläche am Hubende und die je Fläche entzogene Wärme; diese wird mit dem Flächenanteil der beiden Sättel am Umfang auf den Querschnitt verteilt. Gegenüber einer Finite-Differenzen-Rechnung zweier Körper stimmt die Lösung auf 0,01 K überein.
Umformwärme je Drehlage: ΔT = β · k_f · Δφ / (η_F · ρ · c) mit β = 0,9 (Taylor-Quinney) und dem Formänderungswirkungsgrad η_F, gleichmäßig im Querschnitt verteilt. In Wirklichkeit ist die Umformung im Kern größer (Schmiedekreuz), die Kernerwärmung also eher höher.
Luft: Konvektion h und Strahlung ε gegen 20 °C.

Stichplan: QForm-Export
Hubtabelle für die Registerkarte „Blows“ im Format der QForm-Importvorlage (Blatt „Table“; Spalten Blow, CoolingInAir, CoolingInTool, Displacement, RotationAngle; zweite Zeile Einheiten #, s, s, mm, degree). QForm führt alle Werte vor dem jeweiligen Hub aus. CoolingInAir: Transport vor dem ersten Hub, Pause zwischen den Stichen vor dem ersten Hub jedes weiteren Stichs, sonst 0. CoolingInTool: Nebenzeit je Hub vor jedem Hub. Die Verweilzeit unter Last ist nicht Teil der Tabelle; QForm rechnet den Hub selbst. Displacement: Vorschub vor dem ersten Hub jedes neuen Bisses (negativ wie in der Vorlage, d. h. zum Ursprung), optional Rückfahrt an den Stichanfang im ersten Hub des Folgestichs. RotationAngle: Drehschritt vor jedem Hub außer dem ersten Hub eines Stichs. Endabstand der Sättel (= Maß des Stichs) und Hubgeschwindigkeit werden in QForm eingestellt; die Werte stehen in der mitgespeicherten Textdatei.

Stichplan: Fließspannung und Kraft
Sellars-Tegart: σ = 1/α · arsinh[(Z/A)^(1/n)], Z = φ̇ · exp(Q/RT). Hensel-Spittel: σ = A · e^(m1·T) · T^m9 · φ^m2 · e^(m4/φ) · (1+φ)^(m5·T) · e^(m7·φ) · φ̇^m3 · φ̇^(m8·T) mit T in °C, gemittelt über den Umformgrad des Stichs (kumuliert, ohne Entfestigung zwischen den Stichen). Die Vorlagen sind grobe stationäre Richtwerte; die Hensel-Spittel-Startwerte sind eine Anpassung an das Sellars-Tegart-Modell. Umformgeschwindigkeit φ̇ = v / h_mittel.
Presskraft am Hubende: F = k_f · s_B · b_K · Q mit Q = 1,155 · max(0,75 + 0,25 · h/s_B ; 1 + μ · s_B/(4h)), also Hill für schlanke und Siebel für flache Umformzonen.

Stichplan: Nicht enthalten
Erwärmung der Sättel über viele Hübe, Stirnflächen, Breitung beim Vierkant, Abkühlung an der Auflage, Gefügeentwicklung. Die Werte sind Abschätzungen für die Planung; die Parameter h_K und η_F lassen sich mit einer Messung (Thermoelement oder Pyrometer) an der eigenen Presse kalibrieren.

QForm-Simulation
Der Reiter baut den Stichplan über die Python-API von QForm (QFormAPI.py im QForm-Verzeichnis) als 3D-Projekt auf und rechnet ihn Stich für Stich in einem eigenen Prozess. Je Stich eine Umformoperation; ab Stich 2 werden Werkstück und Sättel aus dem Vorstich geerbt und das Werkstück an seiner Ist-Geometrie für den ersten Biss neu positioniert. Untersattel fester Antrieb, Obersattel hydraulische Presse (Nennkraft wählbar, Hubgeschwindigkeit aus dem Stichplan). Stoppbedingung: Abstand der Sättel = Maß des Stichs. Hubtabelle je Stich aus dem Stichplan: vor dem ersten Hub Transport (Stich 1 und nach einer Wiedererwärmung) bzw. Pause an Luft, vor jedem Hub die Nebenzeit auf dem Untersattel, je Biss alle Drehlagen. Bissfolge „vom Manipulator weg“: erster Biss am gegriffenen Ende, Vorschub s·(1 + (λ−1)·(1 − f)). Bissfolge „zum Manipulator hin“: erster Biss am freien Ende, Vorschub s·(1 + (λ−1)·f). Dabei ist f der Anteil der Längung eines Bisses, der die Greifseite verschiebt, weil der Manipulator in QForm nachgibt; mit f wird jeder Biss so vorgeschoben, dass er an den vorigen anschließt (f = 0: starrer Manipulator, Vorschub s·λ bzw. s wie im Stichplan; im Test mit Rückstellung und Widerstand 1,5 etwa 0,5). In beiden Fällen ragt das Ende, an dem die Bissfolge beginnt, eine Bisslänge unter den Sattel, und die ganze Länge wird geschmiedet wie im Stichplan, auch die Greifzone des Manipulators. Bezug ist die ebene Arbeitsfläche des Untersattels ohne Kantenradien (aus dem Netz des Sattels bestimmt); die Sattelbreite im Stichplan sollte dieser Breite entsprechen, sonst bleiben zwischen den Bissen ungeschmiedete Bereiche. Ausgewertet werden je Stich Länge, Querschnitt in der Mitte, Profil über der Länge und die größte Presskraft je Hub. Wiedererwärmung: der Stich startet mit homogener Temperatur statt mit dem geerbten Temperaturfeld. Reibung und Wärmeübergang zum Werkzeug kommen aus dem gewählten Schmierstoff der QForm-Datenbank; Fließspannung aus dem Werkstoff in QForm. Nicht übertragen werden die Verweilzeit unter Last, die Kontaktleitwerte, η_F und μ des Rechners."""

# ---------------------------------------------------------------- Farben
C = dict(
    paper="#EDF0F2", panel="#E1E6EA", field="#FAFBFC", ink="#1C2833", muted="#5A6774",
    line="#C6CDD3", line_strong="#8F9BA6",
    cold="#2D6AA3", cold_soft="#D6E2EE", heat="#D24E17", heat2="#EE9B3E", heat3="#A92A12",
    warn="#8A5200", warn_bg="#FBEFD6", err="#B3261E", err_bg="#FBE6E4", info_bg="#DCE7F1",
    al0="#6E7B87", al1="#EEF2F5", al2="#B3BEC8", al_flat="#C5CED6", al_face="#DCE3E9",
)

# ---------------------------------------------------------------- Geometrie
SQ2 = math.sqrt(2.0)
OCT_K = 2.0 * (SQ2 - 1.0)
OCT_R = SQ2 - 1.0 - math.pi / 8.0


def area(shape, d, r=0.0):
    r = r or 0.0
    if shape == "rund":
        return math.pi * d * d / 4.0
    if shape == "vier":
        return d * d - (4.0 - math.pi) * r * r
    return OCT_K * d * d - 8.0 * OCT_R * r * r


def dim_from_area(shape, A, r=0.0):
    r = r or 0.0
    if shape == "rund":
        return math.sqrt(4.0 * A / math.pi)
    if shape == "vier":
        return math.sqrt(A + (4.0 - math.pi) * r * r)
    return math.sqrt((A + 8.0 * OCT_R * r * r) / OCT_K)


def perimeter_sharp(shape, d):
    if shape == "rund":
        return math.pi * d
    if shape == "vier":
        return 4.0 * d
    return 8.0 * d * (SQ2 - 1.0)


# ---------------------------------------------------------------- Wärmeleitung
SIGMA = 5.670374419e-8


class Slab:
    """1D-Wärmeleitung (Platte oder Zylinder), implizit, dimensionslose Temperatur theta."""

    def __init__(self, R, cyl, N, k, rc, dt):
        self.N = N
        dx = R / N
        V = []
        for i in range(N + 1):
            xi = max(0.0, (i - 0.5) * dx)
            xo = min(R, (i + 0.5) * dx)
            V.append((xo * xo - xi * xi) / 2.0 if cyl else (xo - xi))
        self.V, self.Vt = V, sum(V)
        g = [k * ((i + 0.5) * dx if cyl else 1.0) / dx for i in range(N)]
        self.As = R if cyl else 1.0
        self.Cap = [rc * v / dt for v in V]
        b = [self.Cap[i] + (g[i - 1] if i > 0 else 0.0) + (g[i] if i < N else 0.0) for i in range(N + 1)]
        # Vorwärtselimination ist bis auf den Randknoten unabhängig von h -> einmal vorab
        self.w = [0.0] * (N + 1)
        self.bp = [0.0] * (N + 1)
        self.bp[0] = b[0]
        for i in range(1, N + 1):
            self.w[i] = -g[i - 1] / self.bp[i - 1]
            if i < N:
                self.bp[i] = b[i] + self.w[i] * g[i - 1]
        self.bN = b[N] + self.w[N] * g[N - 1]
        self.g = g
        self.th = [1.0] * (N + 1)

    def step(self, h):
        N, th, Cap, w, g, bp = self.N, self.th, self.Cap, self.w, self.g, self.bp
        d = [Cap[0] * th[0]]
        for i in range(1, N + 1):
            d.append(Cap[i] * th[i] - w[i] * d[i - 1])
        th[N] = d[N] / (self.bN + h * self.As)
        for i in range(N - 1, -1, -1):
            th[i] = (d[i] + g[i] * th[i + 1]) / bp[i]

    def mean(self):
        return sum(v * t for v, t in zip(self.V, self.th)) / self.Vt


def simulate_heating(shape, d, L, rho, cp, k, T0, Tf, Tz, hc, eps, N=30):
    """d, L in m; Temperaturen in °C. Rückgabe: Verläufe und Kennwerte."""
    rc = rho * cp
    is_sq = shape == "vier"
    TfK = Tf + 273.15
    dTh = T0 - Tf

    def h_of(Ts):
        K = Ts + 273.15
        return hc + eps * SIGMA * (K * K + TfK * TfK) * (K + TfK)

    h_max = h_of(Tf)
    A = area(shape, d)
    P = perimeter_sharp(shape, d)
    VoA = A * L / (P * L + 2.0 * A)
    tau = rc * VoA / h_max
    tcond = min(d / 2.0, L / 2.0) ** 2 / (k / rc)
    dt = max(0.02, min(tau / 600.0, max(tcond / 40.0, tau / 6000.0)))
    lat = Slab(d / 2.0, not is_sq, N, k, rc, dt)
    ax = Slab(L / 2.0, False, N, k, rc, dt)
    thZ = (Tz - Tf) / dTh

    def state():
        l0, lN = lat.th[0], lat.th[N]
        lc = l0 * l0 if is_sq else l0
        ls = lN * l0 if is_sq else lN
        le = lN * lN if is_sq else lN
        ac, as_ = ax.th[0], ax.th[N]
        return lc * ac, max(ls * ac, lc * as_), ls * ac, le * as_

    t = 0.0
    core, surf, mid, edge = state()
    ts, cs, ss, es = [0.0], [core], [surf], [edge]
    t_reach = t_surf = mean_reach = None
    max_dt, t_max_dt = 0.0, 0.0
    t_stop, t_limit, steps = 100 * 3600.0, 100 * 3600.0, 0
    while t < t_stop and steps < 300000:
        h = h_of(Tf + dTh * mid)
        lat.step(h)
        ax.step(h)
        t += dt
        steps += 1
        core, surf, mid, edge = state()
        dT = abs(dTh) * (core - edge)
        if dT > max_dt:
            max_dt, t_max_dt = dT, t
        if t_surf is None and surf <= thZ:
            t_surf = t
        if t_reach is None and core <= thZ:
            t_reach = t
            lm = lat.mean()
            mean_reach = (lm * lm if is_sq else lm) * ax.mean()
            t_stop = min(t_limit, t * 1.3)
        ts.append(t)
        cs.append(core)
        ss.append(surf)
        es.append(edge)
    # ausdünnen
    n = len(ts)
    stride = max(1, n // 2500)
    idx = list(range(0, n, stride))
    if idx[-1] != n - 1:
        idx.append(n - 1)

    def conv(arr):
        return [Tf + dTh * arr[i] for i in idx]

    return dict(t=[ts[i] for i in idx], core=conv(cs), surf=conv(ss), edge=conv(es),
                t_reach=t_reach, t_surf=t_surf, max_dt=max_dt, t_max_dt=t_max_dt,
                tau=tau, VoA=VoA, h_max=h_max, Bi=h_max * VoA / k,
                T_mean=None if mean_reach is None else Tf + dTh * mean_reach)


# ---------------------------------------------------------------- Formatierung
def nf(x, d=0):
    if x is None or not math.isfinite(x):
        return "–"
    s = f"{abs(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < 0 and round(abs(x), d) != 0 else "") + s


def fdim(x):
    return nf(x, 0 if abs(x - round(x)) < 0.05 else 1)


def fmt_time(s):
    if s is None or not math.isfinite(s):
        return "–"
    if s < 90:
        return nf(s, 0) + " s"
    m = int(round(s / 60.0))
    if m < 60:
        return f"{m} min"
    return f"{m // 60} h {m % 60:02d} min"


def geo_label(shape, d, L):
    S = SHAPES[shape]
    return f"{S['name']} {S['pre']} {fdim(d)} × {nf(L, 0)} mm"


def parse(s):
    s = (s or "").strip().replace(" ", "").replace(",", ".")
    if not s:
        return float("nan")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def nice_step(rng, n):
    raw = rng / n
    p = 10 ** math.floor(math.log10(raw))
    f = raw / p
    return (1 if f < 1.5 else 2 if f < 3 else 5 if f < 7 else 10) * p


def mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(a[j] + (b[j] - a[j]) * t) for j in range(3))


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def settings_path():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "FreiformschmiedenRechner", "einstellungen.json")


# ---------------------------------------------------------------- Projektdateien
PROJ_EXT = ".ffproj"
PROJ_FORMAT = "Freiformschmieden-Rechner Projekt"
PROJ_FORMAT_VERSION = 1
PROJ_TYPES = [("Freiformschmieden-Projekt", "*" + PROJ_EXT), ("Alle Dateien", "*.*")]
UI_KEYS = ("planView", "planParams",                       # Ansicht und Exportoptionen, keine Projektdaten
           "qfScope", "qfFormat", "qfDec", "qfSign", "qfReturn", "qsLast")
SESSION_META = ("tab", "passes", "projectPath", "projectDirty", "recent")
PROJ_NOTE = ("Projektdatei des Freiformschmieden-Rechners. Beim Öffnen werden die Eingaben geladen und alle "
             "Ergebnisse neu berechnet; der Abschnitt 'ausgaben' dokumentiert die Ergebnisse zum Zeitpunkt des "
             "Speicherns (Längen in mm, Temperaturen in °C, Zeiten in s, Kräfte in MN).")


def clean_json(x, nd=4):
    """Für JSON: Gleitkommazahlen runden, nicht-endliche Werte als null."""
    if isinstance(x, bool) or x is None or isinstance(x, (int, str)):
        return x
    if isinstance(x, float):
        return round(x, nd) if math.isfinite(x) else None
    if isinstance(x, dict):
        return {str(k): clean_json(v, nd) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [clean_json(v, nd) for v in x]
    return str(x)


def dump_project_json(data):
    """JSON mit Einrückung, Zahlenlisten aber einzeilig (gut lesbar und kompakt)."""
    text = json.dumps(data, ensure_ascii=False, indent=1, allow_nan=False)

    def compact(m):
        parts = [p.strip() for p in m.group(1).split(",") if p.strip()]
        return "[" + ", ".join(parts) + "]"
    return re.sub(r"\[([^\[\]{}\"]*)\]", compact, text)


def count_diffs(a, b):
    """Anzahl abweichender Einträge zwischen gespeicherten und neu berechneten Ausgaben (ohne Zeitreihen)."""
    if isinstance(a, dict) and isinstance(b, dict):
        return sum(count_diffs(a.get(k), b.get(k)) for k in set(a) | set(b) if k not in ("verlauf", "kontakte"))
    if isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            return max(1, abs(len(a) - len(b)))
        return sum(count_diffs(x, y) for x, y in zip(a, b))
    num = (int, float)
    if isinstance(a, num) and isinstance(b, num) and not isinstance(a, bool) and not isinstance(b, bool):
        return 0 if abs(a - b) <= max(1e-3 * abs(b), 0.02) else 1
    return 0 if a == b else 1


def decimate(t, y, buckets):
    """Zeitreihe ausdünnen; je Abschnitt bleiben Minimum und Maximum von y erhalten."""
    n = len(t)
    if n <= 2 * buckets:
        return list(range(n))
    keep = {0, n - 1}
    size = n / buckets
    for k in range(buckets):
        lo, hi = int(k * size), min(n, int((k + 1) * size))
        if hi > lo:
            seg = range(lo, hi)
            keep.add(min(seg, key=lambda i: y[i]))
            keep.add(max(seg, key=lambda i: y[i]))
    return sorted(keep)


def short_path(p, n=48):
    p = os.path.normpath(p)
    return p if len(p) <= n else "…" + p[-(n - 1):]


# ---------------------------------------------------------------- Scrollbarer Container
class ScrollFrame(tk.Frame):
    """Füllt den verfügbaren Platz; wird nur bei zu kleinem Fenster vertikal scrollbar."""

    def __init__(self, master, bg):
        super().__init__(master, bg=bg)
        self.cv = tk.Canvas(self, bg=bg, highlightthickness=0, borderwidth=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.cv.yview)
        self.cv.configure(yscrollcommand=self.sb.set)
        self.inner = tk.Frame(self.cv, bg=bg)
        self.win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.grid(row=0, column=0, sticky="nsew")
        self.sb.grid(row=0, column=1, sticky="ns")
        self.sb.grid_remove()
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.cv.bind("<Configure>", lambda e: self._fit())
        self.inner.bind("<Configure>", lambda e: self._fit())
        self.bind("<Enter>", lambda e: self.bind_all("<MouseWheel>", self._wheel))
        self.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))
        self._busy = False

    def _fit(self):
        if self._busy:
            return
        self._busy = True
        try:
            w, h = self.cv.winfo_width(), self.cv.winfo_height()
            need = self.inner.winfo_reqheight()
            hh = max(h, need)
            self.cv.itemconfigure(self.win, width=w, height=hh)
            self.cv.configure(scrollregion=(0, 0, w, hh))
            if need > h + 2:
                self.sb.grid()
            else:
                self.sb.grid_remove()
                self.cv.yview_moveto(0)
        finally:
            self._busy = False

    def refit(self):
        """Nach Inhaltsänderungen aufrufen: Tk meldet geänderte Wunschgrößen nicht als Ereignis."""
        self.after_idle(self._fit)

    def _wheel(self, ev):
        if self.inner.winfo_reqheight() > self.cv.winfo_height() + 2:
            self.cv.yview_scroll(int(-ev.delta / 120) or (-1 if ev.delta > 0 else 1), "units")


class HScroll(tk.Frame):
    """Horizontal scrollbarer Container; Scrollleiste nur bei Bedarf."""

    def __init__(self, master, bg):
        super().__init__(master, bg=bg)
        self.cv = tk.Canvas(self, bg=bg, highlightthickness=0, borderwidth=0, height=10)
        self.sb = ttk.Scrollbar(self, orient="horizontal", command=self.cv.xview)
        self.cv.configure(xscrollcommand=self.sb.set)
        self.inner = tk.Frame(self.cv, bg=bg)
        self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.pack(fill="x")
        self.inner.bind("<Configure>", lambda e: self._fit())
        self.cv.bind("<Configure>", lambda e: self._fit())

    def _fit(self):
        w, h = self.inner.winfo_reqwidth(), self.inner.winfo_reqheight()
        if int(self.cv.cget("height")) != h:
            self.cv.configure(height=h)
        self.cv.configure(scrollregion=(0, 0, w, h))
        if w > self.cv.winfo_width() + 2:
            if not self.sb.winfo_manager():
                self.sb.pack(fill="x")
        elif self.sb.winfo_manager():
            self.sb.pack_forget()
            self.cv.xview_moveto(0)


# ---------------------------------------------------------------- Anwendung
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.configure(bg=C["paper"])
        self.sc = max(1.0, self.winfo_fpixels("1i") / 96.0)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        # nahezu bildschirmfüllend, aber ohne Maximieren (zuverlässig auf allen Systemen)
        w, h = max(int(960 * self.sc), sw - 40), max(int(620 * self.sc), sh - 110)
        self.geometry(f"{min(w, int(1700 * self.sc))}x{min(h, int(1100 * self.sc))}+10+10")
        self.minsize(int(960 * self.sc), int(620 * self.sc))
        self._init_fonts()
        self._init_style()
        self._icons()
        self.v = {}
        self._filling = False
        self._after_heat = None
        self._after_save = None
        self.forge = None
        self.heat = None
        self.plan = None
        self.plan_rows = []
        self.plan_sel = 0
        self._after_plan = None
        self.project_path = None
        self.dirty = False
        self.recent = []
        self._saved_at = None
        self._build_menu()
        self._build_layout()
        self._load()
        self._bind_keys()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._sim_init()
        self.update_all()

    # ---------- Grundeinstellungen
    def _init_fonts(self):
        fams = set(tkfont.families(self))

        def pick(*cands):
            for f in cands:
                if f in fams:
                    return f
            return tkfont.nametofont("TkDefaultFont").actual("family")

        base = pick("Segoe UI", "Barlow", "DejaVu Sans", "Helvetica")
        semi = pick("Segoe UI Semibold", "Barlow SemiBold")
        semi_w = "normal" if semi in ("Segoe UI Semibold", "Barlow SemiBold") else "bold"
        if semi not in ("Segoe UI Semibold", "Barlow SemiBold"):
            semi = base
        self.f_base = tkfont.Font(family=base, size=10)
        self.f_small = tkfont.Font(family=base, size=9)
        self.f_sub = tkfont.Font(family=semi, size=11, weight=semi_w)
        self.f_h = tkfont.Font(family=semi, size=12, weight=semi_w)
        self.f_title = tkfont.Font(family=semi, size=17, weight=semi_w)
        self.f_big = tkfont.Font(family=semi, size=30, weight=semi_w)
        self.f_val = tkfont.Font(family=semi, size=10, weight=semi_w)
        self.f_cv = tkfont.Font(family=base, size=10)
        self.f_cv_b = tkfont.Font(family=semi, size=10, weight=semi_w)
        tkfont.nametofont("TkDefaultFont").configure(family=base, size=10)
        tkfont.nametofont("TkTextFont").configure(family=base, size=10)

    def _init_style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        p = int(4 * self.sc)
        st.configure(".", background=C["paper"], foreground=C["ink"], font=self.f_base,
                     bordercolor=C["line_strong"], lightcolor=C["field"], darkcolor=C["field"],
                     focuscolor=C["cold"])
        st.configure("TFrame", background=C["paper"])
        st.configure("Panel.TFrame", background=C["panel"])
        st.configure("TLabel", background=C["paper"], foreground=C["ink"])
        st.configure("Panel.TLabel", background=C["panel"])
        st.configure("TEntry", fieldbackground=C["field"], padding=(p, p - 1), bordercolor=C["line_strong"],
                     lightcolor=C["field"], darkcolor=C["field"])
        st.map("TEntry", bordercolor=[("focus", C["cold"])], lightcolor=[("focus", C["cold"])])
        st.configure("Calc.TEntry", fieldbackground=C["panel"], foreground=C["muted"], bordercolor=C["line"],
                     lightcolor=C["panel"], darkcolor=C["panel"])
        st.configure("Bad.TEntry", fieldbackground=C["err_bg"], bordercolor=C["err"])
        st.map("Bad.TEntry", bordercolor=[("focus", C["err"])])
        st.configure("TCombobox", fieldbackground=C["field"], padding=(p, p - 1), arrowcolor=C["ink"])
        st.map("TCombobox", fieldbackground=[("readonly", C["field"])], selectbackground=[("readonly", C["field"])],
               selectforeground=[("readonly", C["ink"])])
        self.option_add("*TCombobox*Listbox.font", self.f_base)
        self.option_add("*TCombobox*Listbox.selectBackground", C["cold"])
        st.configure("Seg.Toolbutton", background=C["field"], foreground=C["ink"], padding=(int(8 * self.sc), p),
                     bordercolor=C["line_strong"], relief="solid", borderwidth=1, anchor="center")
        st.map("Seg.Toolbutton",
               background=[("selected", C["ink"]), ("active", C["cold_soft"])],
               foreground=[("selected", C["paper"])])
        st.configure("TNotebook", background=C["paper"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        st.configure("TNotebook.Tab", background=C["paper"], foreground=C["muted"], font=self.f_h,
                     padding=(int(14 * self.sc), int(7 * self.sc)), borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", C["paper"])], foreground=[("selected", C["ink"])],
               expand=[("selected", (0, 0, 0, 0))])
        st.configure("TButton", padding=(int(10 * self.sc), p), background=C["field"])
        st.map("TButton", background=[("active", C["cold_soft"])])

    def _icons(self):
        s = int(16 * self.sc)

        def icon(shape, color):
            img = tk.PhotoImage(width=s, height=s)
            c = (s - 1) / 2.0
            R = s * 0.40
            for y in range(s):
                for x in range(s):
                    px, py = x - c, y - c
                    if shape == "rund":
                        dist = math.hypot(px, py) - R
                    elif shape == "vier":
                        dist = max(abs(px), abs(py)) - R
                    else:
                        dist = max(abs(px), abs(py), (abs(px) + abs(py)) / SQ2) - R
                    if abs(dist) <= 0.85 * self.sc:
                        img.put(color, (x, y))
            return img

        self.ic = {k: (icon(k, C["ink"]), icon(k, C["paper"])) for k in SHAPES}
        # Fenstersymbol
        n = 32
        win = tk.PhotoImage(width=n, height=n)
        c = (n - 1) / 2.0
        for y in range(n):
            for x in range(n):
                px, py = x - c, y - c
                do = max(abs(px), abs(py), (abs(px) + abs(py)) / SQ2) - 14.5
                dc = math.hypot(px, py) - 7.0
                if abs(dc) <= 1.1:
                    win.put(C["heat"], (x, y))
                elif abs(do) <= 1.1:
                    win.put(C["ink"], (x, y))
        self._win_icon = win
        ico = resource_path("icon.ico")
        if sys.platform == "win32" and os.path.exists(ico):
            try:
                self.iconbitmap(default=ico)
                return
            except tk.TclError:
                pass
        self.iconphoto(True, win)

    # ---------- Menü
    def _build_menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="Neues Projekt", accelerator="Strg+N", command=self.new_project)
        f.add_command(label="Projekt öffnen …", accelerator="Strg+O", command=self.open_project_dialog)
        self.m_recent = tk.Menu(f, tearoff=0)
        f.add_cascade(label="Zuletzt geöffnet", menu=self.m_recent)
        f.add_command(label="Projekt speichern", accelerator="Strg+S", command=self.save_project)
        f.add_command(label="Projekt speichern unter …", accelerator="Strg+Umschalt+S",
                      command=self.save_project_as)
        f.add_separator()
        f.add_command(label="Stichplan als Excel-Datei speichern …", command=self.export_plan_xlsx)
        f.add_command(label="Stichplan als CSV speichern …", command=self.export_plan_csv)
        f.add_command(label="Stichplan für QForm exportieren …", command=self.export_qform)
        f.add_command(label="Aufheizkurve als CSV speichern …", command=self.export_csv)
        f.add_command(label="Ergebnisse in die Zwischenablage kopieren", command=self.copy_summary)
        f.add_separator()
        f.add_command(label="Beenden", command=self._on_close)
        m.add_cascade(label="Datei", menu=f)
        h = tk.Menu(m, tearoff=0)
        h.add_command(label="Rechenmodell und Annahmen", command=self.show_model)
        h.add_command(label="Über", command=lambda: messagebox.showinfo(
            "Über", f"{APP_TITLE}\nVersion {VERSION}\n\nEndlänge nach Volumenkonstanz und Aufheizzeit "
                    "nach instationärer Wärmeleitung mit Konvektion und Strahlung.", parent=self))
        m.add_cascade(label="Hilfe", menu=h)
        self.config(menu=m)

    # ---------- Hilfen für Eingabefelder
    def var(self, key):
        if key not in self.v:
            sv = tk.StringVar(self)
            sv.trace_add("write", lambda *_a, k=key: self._changed(k))
            self.v[key] = sv
        return self.v[key]

    def field(self, parent, row, label, key, unit, bg, width=9):
        lbl = tk.Label(parent, text=label, bg=bg, fg=C["ink"], font=self.f_base, anchor="w")
        lbl.grid(row=row, column=0, sticky="w", pady=int(3 * self.sc))
        e = ttk.Entry(parent, textvariable=self.var(key), width=width, justify="right")
        e.grid(row=row, column=1, sticky="e", padx=(int(8 * self.sc), int(4 * self.sc)), pady=int(3 * self.sc))
        u = tk.Label(parent, text=unit, bg=bg, fg=C["muted"], font=self.f_small, anchor="w", width=8)
        u.grid(row=row, column=2, sticky="w")
        e._lbl = lbl
        return e

    def seg(self, parent, key, options, bg, icons=False):
        fr = tk.Frame(parent, bg=bg)
        for i, (val, text) in enumerate(options):
            kw = {}
            if icons:
                kw = dict(image=(self.ic[val][0], "selected", self.ic[val][1]), compound="left")
            rb = ttk.Radiobutton(fr, text=" " + text, value=val, variable=self.var(key), style="Seg.Toolbutton", **kw)
            rb.grid(row=0, column=i, sticky="ew")
            fr.columnconfigure(i, weight=1, uniform="seg")
        return fr

    def hint(self, parent, text, bg, wrap=290):
        return tk.Label(parent, text=text, bg=bg, fg=C["muted"], font=self.f_small, justify="left",
                        anchor="w", wraplength=int(wrap * self.sc))

    def mark(self, entry, bad):
        entry.configure(style="Bad.TEntry" if bad else "TEntry")
        return bad

    # ---------- Aufbau
    def _build_layout(self):
        sc = self.sc
        head = tk.Frame(self, bg=C["paper"])
        head.pack(fill="x", padx=int(20 * sc), pady=(int(14 * sc), 0))
        tk.Label(head, image=self._win_icon, bg=C["paper"]).pack(side="left", padx=(0, int(10 * sc)))
        tt = tk.Frame(head, bg=C["paper"])
        tt.pack(side="left")
        tk.Label(tt, text="Freiformschmieden Aluminium", font=self.f_title, bg=C["paper"], fg=C["ink"]).pack(anchor="w")
        tk.Label(tt, text="Vorberechnung von Endlänge, Aufheizzeit und Stichplan", font=self.f_base, bg=C["paper"],
                 fg=C["muted"]).pack(anchor="w")
        pb = tk.Frame(head, bg=C["paper"])
        pb.pack(side="right")
        ttk.Button(pb, text="Speichern", command=self.save_project).pack(side="right")
        ttk.Button(pb, text="Öffnen …", command=self.open_project_dialog).pack(side="right", padx=(0, int(6 * sc)))
        ttk.Button(pb, text="Neu", command=self.new_project).pack(side="right", padx=(0, int(6 * sc)))
        pi = tk.Frame(pb, bg=C["paper"])
        pi.pack(side="right", padx=(0, int(16 * sc)))
        self.lb_proj = tk.Label(pi, text="", font=self.f_sub, bg=C["paper"], fg=C["ink"], anchor="e")
        self.lb_proj.pack(anchor="e")
        self.lb_proj_state = tk.Label(pi, text="", font=self.f_small, bg=C["paper"], fg=C["muted"], anchor="e")
        self.lb_proj_state.pack(anchor="e")
        tk.Frame(self, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x", padx=int(20 * sc), pady=(int(10 * sc), 0))

        body = tk.Frame(self, bg=C["paper"])
        body.pack(fill="both", expand=True, padx=(int(20 * sc), int(20 * sc)), pady=(0, int(12 * sc)))
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # --- Seitenleiste
        side = tk.Frame(body, bg=C["panel"], padx=int(16 * sc), pady=int(10 * sc))
        side.grid(row=0, column=0, sticky="ns")
        bg = C["panel"]
        tk.Label(side, text="Werkstoff", font=self.f_h, bg=bg, fg=C["ink"]).grid(row=0, column=0, columnspan=3, sticky="w")
        self.cb_mat = ttk.Combobox(side, textvariable=self.var("matPreset"), state="readonly",
                                   values=[m["name"] for m in MATS] + [CUSTOM], width=28)
        self.cb_mat.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(int(6 * sc), int(4 * sc)))
        self.cb_mat.bind("<<ComboboxSelected>>", lambda e: self._fill_material())
        self.hint(side, "Richtwerte bei Raumtemperatur. cp steigt bis 500 °C um rund 20 %, "
                        "Mittelwert verwenden.", bg).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, int(4 * sc)))
        tk.Label(side, text="Bezeichnung", bg=bg, font=self.f_base).grid(row=3, column=0, sticky="w", pady=int(3 * sc))
        ttk.Entry(side, textvariable=self.var("matName"), width=18).grid(row=3, column=1, columnspan=2, sticky="ew",
                                                                         padx=(int(8 * sc), 0))
        self.e_rho = self.field(side, 4, "Dichte ρ", "rho", "kg/m³", bg)
        self.e_cp = self.field(side, 5, "Wärmekapazität cp", "cp", "J/(kg K)", bg)
        self.e_lam = self.field(side, 6, "Wärmeleitfähigkeit λ", "lam", "W/(m K)", bg)
        self.lb_alpha = tk.Label(side, text="", bg=bg, fg=C["muted"], font=self.f_small, anchor="w")
        self.lb_alpha.grid(row=7, column=0, columnspan=3, sticky="w", pady=(int(2 * sc), 0))

        tk.Frame(side, bg=C["line_strong"], height=1).grid(row=8, column=0, columnspan=3, sticky="ew", pady=int(10 * sc))
        tk.Label(side, text="Rohblock", font=self.f_h, bg=bg).grid(row=9, column=0, columnspan=3, sticky="w")
        self.seg(side, "blkShape", [(k, S["name"]) for k, S in SHAPES.items()], bg, icons=True).grid(
            row=10, column=0, columnspan=3, sticky="ew", pady=(int(6 * sc), int(6 * sc)))
        self.e_blkD = self.field(side, 11, "Durchmesser D₀", "blkD", "mm", bg)
        self.e_blkR = self.field(side, 12, "Kantenradius r₀", "blkR", "mm", bg)
        self.e_blkL = self.field(side, 13, "Länge L₀", "blkL", "mm", bg)
        self.blk_info = tk.Frame(side, bg=bg)
        self.blk_info.grid(row=14, column=0, columnspan=3, sticky="ew", pady=(int(8 * sc), 0))
        side.columnconfigure(0, weight=1)

        # --- Hauptbereich mit Reitern
        main = tk.Frame(body, bg=C["paper"])
        main.grid(row=0, column=1, sticky="nsew", padx=(int(22 * sc), 0))
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        self.nb = ttk.Notebook(main)
        self.nb.grid(row=0, column=0, sticky="nsew", pady=(int(4 * sc), 0))
        self.sf_f = ScrollFrame(self.nb, C["paper"])
        self.sf_h = ScrollFrame(self.nb, C["paper"])
        self.sf_p = ScrollFrame(self.nb, C["paper"])
        self.sf_q = ScrollFrame(self.nb, C["paper"])
        self.tab_f, self.tab_h, self.tab_p = self.sf_f.inner, self.sf_h.inner, self.sf_p.inner
        self.tab_q = self.sf_q.inner
        self.nb.add(self.sf_f, text="Endlänge nach dem Schmieden")
        self.nb.add(self.sf_h, text="Aufheizzeit im Ofen")
        self.nb.add(self.sf_p, text="Stichplan")
        self.nb.add(self.sf_q, text="QForm-Simulation")
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._tab_changed())
        self._build_forge_tab()
        self._build_heat_tab()
        self._build_plan_tab()
        self._build_sim_tab()

    def _section(self, parent, title):
        fr = tk.Frame(parent, bg=C["paper"])
        tk.Frame(fr, bg=C["ink"], height=max(2, int(2 * self.sc))).pack(fill="x")
        tk.Label(fr, text=title, font=self.f_h, bg=C["paper"]).pack(anchor="w", pady=(int(6 * self.sc), int(4 * self.sc)))
        inner = tk.Frame(fr, bg=C["paper"])
        inner.pack(fill="both", expand=True)
        return fr, inner

    def _build_forge_tab(self):
        sc, bg = self.sc, C["paper"]
        t = self.tab_f
        t.columnconfigure(1, weight=1)
        t.rowconfigure(1, weight=1)
        fr, inn = self._section(t, "Schmiedestück")
        fr.grid(row=0, column=0, sticky="nw", padx=(0, int(28 * sc)), pady=(int(14 * sc), 0))
        self.seg(inn, "tgtShape", [(k, S["name"]) for k, S in SHAPES.items()], bg, icons=True).grid(
            row=0, column=0, columnspan=3, sticky="ew", pady=(0, int(6 * sc)))
        self.seg(inn, "tgtMode", [("dim", "Querschnitt vorgeben"), ("len", "Endlänge vorgeben")], bg).grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(0, int(6 * sc)))
        self.e_tgtD = self.field(inn, 2, "Kantenlänge a₁", "tgtD", "mm", bg)
        self.e_tgtL = self.field(inn, 3, "Endlänge L₁", "tgtL", "mm", bg)
        self.e_tgtR = self.field(inn, 4, "Kantenradius r₁", "tgtR", "mm", bg)
        self.e_loss = self.field(inn, 5, "Materialverlust", "loss", "%", bg)
        self.hint(inn, "z. B. Endenabschnitt oder Bearbeitungszugabe, bezogen auf das Volumen des Rohblocks.", bg,
                  wrap=280).grid(row=6, column=0, columnspan=3, sticky="w")
        inn.columnconfigure(0, weight=1)

        fr2, self.res_f = self._section(t, "Ergebnis")
        fr2.grid(row=0, column=1, sticky="new", pady=(int(14 * sc), 0))

        cf = tk.Frame(t, bg=bg)
        cf.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(int(16 * sc), 0))
        tk.Frame(cf, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x")
        hd = tk.Frame(cf, bg=bg)
        hd.pack(fill="x", pady=(int(6 * sc), int(2 * sc)))
        tk.Label(hd, text="Maßstäbliche Darstellung", font=self.f_h, bg=bg).pack(side="left")
        tk.Label(hd, text="Querschnitte und Längen jeweils zueinander maßstäblich", font=self.f_small, fg=C["muted"],
                 bg=bg).pack(side="right")
        self.cv_f = tk.Canvas(cf, bg=bg, highlightthickness=0, height=int(240 * sc))
        self.cv_f.pack(fill="both", expand=True)
        self.cv_f.bind("<Configure>", lambda e: self.draw_forge())

    def _build_heat_tab(self):
        sc, bg = self.sc, C["paper"]
        t = self.tab_h
        t.columnconfigure(2, weight=1)
        t.rowconfigure(2, weight=1)
        fr, inn = self._section(t, "Bauteil und Temperaturen")
        fr.grid(row=0, column=0, sticky="nw", padx=(0, int(22 * sc)), pady=(int(10 * sc), 0))
        self.seg(inn, "heatGeo", [("blk", "Rohblock"), ("tgt", "Schmiedestück")], bg).grid(
            row=0, column=0, columnspan=3, sticky="ew", pady=(0, int(4 * sc)))
        self.lb_geo = self.hint(inn, "", bg, wrap=320)
        self.lb_geo.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, int(6 * sc)))
        self.e_T0 = self.field(inn, 2, "Einsatztemperatur T₀", "T0", "°C", bg)
        self.e_Tf = self.field(inn, 3, "Ofentemperatur", "Tf", "°C", bg)
        self.e_Tz = self.field(inn, 4, "Zieltemperatur Kern", "Tz", "°C", bg)
        inn.columnconfigure(0, weight=1)

        fr2, inn2 = self._section(t, "Ofen und Wärmeübergang")
        fr2.grid(row=0, column=1, sticky="nw", padx=(0, int(26 * sc)), pady=(int(10 * sc), 0))
        self.cb_furn = ttk.Combobox(inn2, textvariable=self.var("furnPreset"), state="readonly",
                                    values=[f["name"] for f in FURN] + [CUSTOM], width=30)
        self.cb_furn.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.cb_furn.bind("<<ComboboxSelected>>", lambda e: self._fill_furnace())
        self.lb_furn = self.hint(inn2, "", bg, wrap=300)
        self.lb_furn.grid(row=1, column=0, columnspan=3, sticky="w", pady=(int(4 * sc), int(4 * sc)))
        self.e_hc = self.field(inn2, 2, "Konvektion h", "hc", "W/(m² K)", bg)
        self.e_eps = self.field(inn2, 3, "Emissionsgrad ε", "eps", "–", bg)
        self.hint(inn2, "Al blank 0,05; gesägt oder oxidiert 0,1 bis 0,2.", bg, wrap=300).grid(
            row=4, column=0, columnspan=3, sticky="w")
        inn2.columnconfigure(0, weight=1)

        fr3, inn3 = self._section(t, "Ergebnis")
        fr3.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(int(8 * sc), 0))
        self.res_h = tk.Frame(inn3, bg=bg, width=int(300 * sc))
        self.res_h.grid(row=0, column=0, sticky="nw", padx=(0, int(18 * sc)))
        self.res_h2 = tk.Frame(inn3, bg=bg)
        self.res_h2.grid(row=0, column=1, sticky="new")
        self._tile_labels = []
        self.res_h2.bind("<Configure>", lambda e: self._tile_wrap(e.width))
        inn3.columnconfigure(1, weight=1)

        cf = tk.Frame(t, bg=bg)
        cf.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(int(4 * sc), 0))
        tk.Frame(cf, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x")
        hd = tk.Frame(cf, bg=bg)
        hd.pack(fill="x", pady=(int(6 * sc), int(2 * sc)))
        self.lb_chart = tk.Label(hd, text="Temperaturverlauf", font=self.f_h, bg=bg)
        self.lb_chart.pack(side="left")
        ttk.Button(hd, text="Als CSV speichern …", command=self.export_csv).pack(side="right")
        lg = tk.Frame(cf, bg=bg)
        lg.pack(fill="x", pady=(0, int(2 * sc)))
        self.lg_labels = []
        for txt, col, wd, dash in (("Kern", C["heat"], 3, None), ("Oberfläche, kälteste Stelle", C["heat2"], 2, None),
                                   ("Kante, wärmste Stelle", C["heat3"], 2, (2, 3))):
            sw = tk.Canvas(lg, width=int(24 * sc), height=int(12 * sc), bg=bg, highlightthickness=0)
            kw = dict(dash=dash) if dash else {}
            sw.create_line(0, int(6 * sc), int(24 * sc), int(6 * sc), fill=col, width=wd * sc, **kw)
            sw.pack(side="left", padx=(0, int(5 * sc)))
            l = tk.Label(lg, text=txt, font=self.f_small, fg=C["muted"], bg=bg)
            l.pack(side="left", padx=(0, int(16 * sc)))
            self.lg_labels.append(l)
        self.cv_h = tk.Canvas(cf, bg=bg, highlightthickness=0, height=int(220 * sc))
        self.cv_h.pack(fill="both", expand=True)
        self.cv_h.bind("<Configure>", lambda e: self.draw_chart())
        self.cv_h.bind("<Motion>", self._hover)
        self.cv_h.bind("<Leave>", lambda e: self._hover_clear())
        self.lb_read = tk.Label(cf, text="Mauszeiger über das Diagramm bewegen, um Werte abzulesen.", font=self.f_base,
                                fg=C["muted"], bg=bg, anchor="w")
        self.lb_read.pack(side="bottom", fill="x", before=self.cv_h)

    # ---------- Laden / Speichern
    def _load(self):
        """Letzte Sitzung wiederherstellen, einschließlich ungespeicherter Änderungen am Projekt."""
        data = dict(DEFAULTS)
        try:
            with open(settings_path(), "r", encoding="utf-8") as fh:
                data.update(json.load(fh))
        except Exception:
            pass
        self._apply(data)
        p = data.get("projectPath")
        self.project_path = p if isinstance(p, str) and p else None
        self.dirty = bool(data.get("projectDirty")) if self.project_path or data.get("projectDirty") else False
        rec = data.get("recent")
        self.recent = [x for x in rec if isinstance(x, str)][:8] if isinstance(rec, list) else []
        self._rebuild_recent_menu()
        self._update_project_ui()

    def _apply(self, data):
        self._filling = True
        if not all(("hs" + k) in data for k in sp_HS_KEYS) or "kfLnA" not in data:
            m = next((m for m in MATS if m["name"] == data.get("matPreset")), MATS[0])
            for k, val in kf_fields(m).items():
                data.setdefault(k, val)
        for k, val in data.items():
            if k in SESSION_META:
                continue
            self.var(k).set(val)
        self._filling = False
        passes = data.get("passes")
        if not isinstance(passes, list):
            passes = DEFAULT_PASSES
        self._set_plan_rows(passes)
        try:
            self.nb.select(int(data.get("tab", 0)))
        except Exception:
            pass

    def _save(self):
        data = {k: sv.get() for k, sv in self.v.items()}
        data["passes"] = self._plan_rows_data()
        try:
            data["tab"] = str(self.nb.index(self.nb.select()))
        except Exception:
            data["tab"] = "0"
        data["projectPath"] = self.project_path or ""
        data["projectDirty"] = self.dirty
        data["recent"] = self.recent
        try:
            p = settings_path()
            os.makedirs(os.path.dirname(p), exist_ok=True)
            with open(p, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _on_close(self):
        if getattr(self, "sim_proc", None) is not None and self.sim_proc.poll() is None:
            if not messagebox.askokcancel("Beenden", "Die QForm-Simulation läuft im Hintergrund weiter und schreibt "
                                                     "ihr Protokoll in den Ausgabeordner. Programm schließen?",
                                          parent=self):
                return
        for job in (self._after_save, self._after_heat, self._after_plan, getattr(self, "_after_sim", None),
                    getattr(self, "_after_simpoll", None)):
            if job:
                try:
                    self.after_cancel(job)
                except tk.TclError:
                    pass
        self._save()
        self.destroy()

    def reset(self):
        d = dict(DEFAULTS)
        d.update(kf_fields(MATS[0]))
        d["passes"] = DEFAULT_PASSES
        self._apply(d)
        self.nb.select(0)
        self.update_all()

    # ================================================================ Projekte
    def _bind_keys(self):
        for seq, fn in (("<Control-n>", self.new_project), ("<Control-N>", self.new_project),
                        ("<Control-o>", self.open_project_dialog), ("<Control-O>", self.open_project_dialog),
                        ("<Control-s>", self.save_project), ("<Control-S>", self.save_project_as)):
            self.bind(seq, lambda e, f=fn: (f(), "break")[1])

    def _project_name(self):
        if not self.project_path:
            return "Unbenanntes Projekt"
        return os.path.splitext(os.path.basename(self.project_path))[0]

    def _project_dir(self):
        for p in ([self.project_path] if self.project_path else []) + self.recent:
            d = os.path.dirname(p)
            if d and os.path.isdir(d):
                return d
        return os.path.expanduser("~")

    def _mark_dirty(self):
        if not self.dirty:
            self.dirty = True
            self._update_project_ui()

    def _update_project_ui(self):
        name = self._project_name()
        self.title(f"{name}{' *' if self.dirty else ''} – Freiformschmieden-Rechner")
        self.lb_proj.configure(text=name)
        if self.dirty:
            txt, col = "Ungespeicherte Änderungen", C["warn"]
        elif self.project_path and self._saved_at:
            t = self._saved_at
            when = f"um {t:%H:%M}" if t.date() == datetime.date.today() else f"am {t:%d.%m.%Y}"
            txt, col = f"Gespeichert {when} in {short_path(os.path.dirname(self.project_path))}", C["muted"]
        elif self.project_path:
            txt, col = short_path(os.path.dirname(self.project_path)), C["muted"]
        else:
            txt, col = "Noch nicht als Projektdatei gespeichert", C["muted"]
        self.lb_proj_state.configure(text=txt, fg=col)

    def _confirm_discard(self, before):
        """Bei ungespeicherten Änderungen nachfragen. True = fortfahren."""
        if not self.dirty:
            return True
        ans = messagebox.askyesnocancel(
            "Ungespeicherte Änderungen",
            f"„{self._project_name()}“ enthält ungespeicherte Änderungen.\n\nVor dem {before} speichern?",
            parent=self)
        if ans is None:
            return False
        if ans:
            return self.save_project()
        return True

    # ---------- Neu / Öffnen / Speichern
    def new_project(self):
        if not self._confirm_discard("Anlegen eines neuen Projekts"):
            return False
        self.reset()
        self.project_path, self.dirty, self._saved_at = None, False, None
        self._update_project_ui()
        self._save()
        return True

    def open_project_dialog(self):
        if not self._confirm_discard("Öffnen eines anderen Projekts"):
            return False
        path = filedialog.askopenfilename(parent=self, title="Projekt öffnen", filetypes=PROJ_TYPES,
                                          initialdir=self._project_dir())
        if not path:
            return False
        return self.open_project(path, confirm=False)

    def open_project(self, path, confirm=True):
        if confirm and not self._confirm_discard("Öffnen eines anderen Projekts"):
            return False
        try:
            with open(path, "r", encoding="utf-8-sig") as fh:
                data = json.load(fh)
            if not isinstance(data, dict) or data.get("format") != PROJ_FORMAT:
                raise ValueError("Die Datei ist keine Projektdatei des Freiformschmieden-Rechners.")
        except (OSError, ValueError) as e:
            if not os.path.exists(path):
                self.recent = [p for p in self.recent if os.path.normcase(p) != os.path.normcase(os.path.abspath(path))]
                self._rebuild_recent_menu()
                msg = "Die Datei wurde nicht gefunden."
            elif isinstance(e, json.JSONDecodeError):
                msg = f"Der Inhalt ist kein gültiges JSON (Zeile {e.lineno}, Spalte {e.colno})."
            else:
                msg = str(e)
            messagebox.showerror("Projekt öffnen", f"{path}\n\n{msg}", parent=self)
            return False
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            diffs = self._apply_project(data)
        finally:
            self.config(cursor="")
        self.project_path = os.path.abspath(path)
        self.dirty, self._saved_at = False, None
        try:
            self._saved_at = datetime.datetime.fromisoformat(str(data.get("gespeichert")))
        except ValueError:
            pass
        self._add_recent(self.project_path)
        self._update_project_ui()
        self._save()
        notes = []
        fv = data.get("formatversion", 1)
        if isinstance(fv, int) and fv > PROJ_FORMAT_VERSION:
            notes.append("Das Projekt wurde mit einer neueren Programmversion gespeichert. Angaben, die diese "
                         "Version nicht kennt, wurden übergangen.")
        if diffs:
            ver = data.get("programmversion", "?")
            when = str(data.get("gespeichert", "")).replace("T", " ")
            notes.append(f"Die Ergebnisse wurden neu berechnet und weichen in {diffs} Werten von den in der Datei "
                         f"gespeicherten ab (gespeichert mit Version {ver}, {when}). Das kommt vor, wenn sich das "
                         "Rechenmodell zwischen Programmversionen geändert hat. Beim nächsten Speichern werden die "
                         "neuen Ergebnisse übernommen.")
        if notes:
            messagebox.showinfo("Projekt geöffnet", "\n\n".join(notes), parent=self)
        return True

    def save_project(self):
        if not self.project_path:
            return self.save_project_as()
        return self._do_save(self.project_path)

    def save_project_as(self):
        if self.project_path:
            init = os.path.basename(self.project_path)
        else:
            mat = "".join(c if c.isalnum() or c in "-_" else "_" for c in self.v["matName"].get().strip()) or "Projekt"
            init = f"Projekt_{mat}_{datetime.date.today():%Y-%m-%d}{PROJ_EXT}"
        path = filedialog.asksaveasfilename(parent=self, title="Projekt speichern unter", defaultextension=PROJ_EXT,
                                            initialfile=init, initialdir=self._project_dir(), filetypes=PROJ_TYPES)
        if not path:
            return False
        return self._do_save(path)

    def _do_save(self, path):
        self.config(cursor="watch")
        self.update_idletasks()
        tmp = path + ".tmp"
        try:
            text = dump_project_json(self._project_dict())
            with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(text)
            os.replace(tmp, path)
        except PermissionError:
            messagebox.showerror("Projekt speichern", f"{path}\n\nDie Datei ist schreibgeschützt oder in einem "
                                                      "anderen Programm geöffnet.", parent=self)
            return False
        except (OSError, ValueError) as e:
            messagebox.showerror("Projekt speichern", f"{path}\n\nDie Datei konnte nicht geschrieben werden:\n{e}",
                                 parent=self)
            return False
        finally:
            self.config(cursor="")
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        self.project_path = os.path.abspath(path)
        self.dirty, self._saved_at = False, datetime.datetime.now()
        self._add_recent(self.project_path)
        self._update_project_ui()
        self._save()
        return True

    # ---------- Zuletzt geöffnet
    def _add_recent(self, path):
        p = os.path.abspath(path)
        self.recent = ([p] + [x for x in self.recent if os.path.normcase(x) != os.path.normcase(p)])[:8]
        self._rebuild_recent_menu()

    def _rebuild_recent_menu(self):
        m = self.m_recent
        m.delete(0, "end")
        if not self.recent:
            m.add_command(label="(noch keine)", state="disabled")
            return
        for i, p in enumerate(self.recent):
            m.add_command(label=f"{i + 1}   {os.path.basename(p)}   ({short_path(os.path.dirname(p), 40)})",
                          command=lambda p=p: self.open_project(p))
        m.add_separator()
        m.add_command(label="Liste leeren", command=self._clear_recent)

    def _clear_recent(self):
        self.recent = []
        self._rebuild_recent_menu()
        self._save()

    # ---------- Inhalt der Projektdatei
    def _ensure_results(self):
        """Alle drei Rechnungen auf den Stand der Eingaben bringen, unabhängig vom sichtbaren Reiter."""
        self.update_all()
        self.run_heat()
        self.run_plan()

    def _project_dict(self):
        self._ensure_results()
        inputs = {k: sv.get() for k, sv in self.v.items() if k not in UI_KEYS}
        view = {k: self.v[k].get() for k in UI_KEYS if k in self.v}
        try:
            view["tab"] = str(self.nb.index(self.nb.select()))
        except tk.TclError:
            view["tab"] = "0"
        return dict(format=PROJ_FORMAT, formatversion=PROJ_FORMAT_VERSION, programmversion=VERSION,
                    gespeichert=datetime.datetime.now().isoformat(timespec="seconds"), hinweis=PROJ_NOTE,
                    eingaben=dict(sorted(inputs.items())), stiche=self._plan_rows_data(), ansicht=view,
                    ausgaben=clean_json(self._collect_outputs()))

    def _apply_project(self, data):
        """Eingaben aus einer Projektdatei übernehmen, neu rechnen; Rückgabe: Zahl abweichender Ergebnisse."""
        known = set(DEFAULTS) | set(kf_fields(MATS[0])) | {"matName"}
        raw = data.get("eingaben") if isinstance(data.get("eingaben"), dict) else {}
        d = dict(DEFAULTS)
        for k, v in raw.items():
            if k in known and k not in UI_KEYS and k != "tab":
                if isinstance(v, bool):
                    continue
                d[k] = fe(v, 6).rstrip("0").rstrip(",") if isinstance(v, (int, float)) else str(v)
        view = data.get("ansicht") if isinstance(data.get("ansicht"), dict) else {}
        for k in UI_KEYS + ("tab",):
            if k in view:
                d[k] = str(view[k])
        passes = []
        for p in data.get("stiche") or []:
            if not isinstance(p, dict):
                continue
            passes.append(dict(shape=p.get("shape") if p.get("shape") in SHAPES else "vier",
                               mode=p.get("mode") if p.get("mode") in ("dim", "len") else "dim",
                               dim=str(p.get("dim", "") or ""), len=str(p.get("len", "") or ""),
                               nrot=str(p.get("nrot", "") or ""), bk=str(p.get("bk", "") or ""),
                               rw=str(p.get("rw", "") or ""), rwT=str(p.get("rwT", "") or "")))
        d["passes"] = passes                                  # fehlende Fließspannungswerte ergänzt _apply
        self.plan_sel = 0
        self._apply(d)
        self._ensure_results()
        stored = data.get("ausgaben")
        if not isinstance(stored, dict):
            return 0
        return count_diffs(stored, clean_json(self._collect_outputs()))

    def _collect_outputs(self):
        """Ergebnisse aller drei Reiter in lesbarer Form (Längen mm, Temperaturen °C, Zeiten s)."""
        out = {}
        f = self.forge
        if f and f.get("ok"):
            e = dict(form=f["shape"], vorgabe="Querschnitt" if f["mode"] == "dim" else "Endlänge",
                     masz_mm=f["d1"], kantenradius_mm=f["r1"], laenge_mm=f["L1"],
                     querschnitt_rohblock_mm2=f["A0"], querschnitt_mm2=f["A1"], volumen_mm3=f["Veff"],
                     materialverlust_prozent=f["loss"], verschmiedungsgrad=f["ratio"], umformgrad=f["phi"])
            if f["mat"]["ok"]:
                e["masse_kg"] = f["mat"]["rho"] * f["Veff"] * 1e-9
            e["hinweise"] = [t for _k, t in f["warns"]]
            out["endlaenge"] = e
        else:
            out["endlaenge"] = dict(fehler=list(f["errs"]) if f else ["nicht berechnet"])
        h = self.heat
        if h:
            res, geo = h["res"], h["geo"]
            a = dict(bauteil=geo["label"], form=geo["shape"], masz_mm=geo["d"], laenge_mm=geo["L"],
                     aufheizzeit_s=res["t_reach"], oberflaeche_erreicht_s=res["t_surf"],
                     max_temperaturdifferenz_K=res["max_dt"], zeitpunkt_max_temperaturdifferenz_s=res["t_max_dt"],
                     biot_zahl=res["Bi"], zeitkonstante_s=res["tau"],
                     waermeuebergang_bei_ofentemperatur_W_m2K=res["h_max"], masse_kg=h["mass"])
            if res["T_mean"] is not None:
                a["waermemenge_kWh"] = h["mass"] * h["mat"]["cp"] * (res["T_mean"] - h["T0"]) / 3.6e6
            idx = decimate(res["t"], res["surf"], 150)
            a["verlauf"] = dict(t_s=[round(res["t"][i], 1) for i in idx],
                                kern_C=[round(res["core"][i], 2) for i in idx],
                                oberflaeche_C=[round(res["surf"][i], 2) for i in idx],
                                kante_C=[round(res["edge"][i], 2) for i in idx])
            out["aufheizen"] = a
        else:
            out["aufheizen"] = dict(fehler=["Eingaben unvollständig, nicht berechnet"])
        P = self.plan
        if P and P["res"] and P["res"]["rows"]:
            D = self._plan_export_data()
            rows = []
            for r in D["rows"]:
                rows.append(dict(
                    nr=r["nr"], form=r["shape"], masz_mm=r["d"], laenge_mm=r["L"], querschnitt_mm2=r["A"],
                    verschmiedungsgrad=r["ratio"], umformgrad=r["phi"], umformgrad_gesamt=r["phi_cum"],
                    drehlagen=r["n_rot"], drehschritt_grad=r["rot_step"], bisse=r["n_b"], huebe=r["strokes"],
                    bisslaenge_mm=r["sB"], letzter_biss_mm=r["s_last"], vorschub_mm=r["feed"],
                    kontaktbreite_mm=r["bc"], kontaktflaeche_je_sattel_mm2=r["A_K"], hubweg_mm=r["dh"],
                    kontaktzeit_s=r["t_c"], umformgeschwindigkeit_1_s=r["phid"], kf_MPa=r["kf"],
                    presskraft_MN=r["F"], stichdauer_s=r["t_pass"], zeit_stichende_s=r["t_end"],
                    dT_umformwaerme_K=r["dT_def"], dT_sattelkontakt_K=-r["dT_die"], dT_mitteltemperatur_K=r["dT_mean"],
                    T_kern_C=r["T_core"], T_oberflaeche_C=r["T_surf"], T_mittel_C=r["T_mean"],
                    T_kontaktzone_min_C=r["T_contact"]))
            st, s = D["start"], D["series"]
            idx = decimate(s["t"], s["surf"], 300)
            out["stichplan"] = dict(
                start=dict(T_kern_C=st[0], T_oberflaeche_C=st[1], T_mittel_C=st[2]),
                stiche=rows,
                hinweise=[dict(art=k, text=t) for k, t in D["notes"]],
                verlauf=dict(t_s=[round(s["t"][i], 2) for i in idx], kern_C=[round(s["core"][i], 2) for i in idx],
                             oberflaeche_C=[round(s["surf"][i], 2) for i in idx]),
                kontakte=[[round(t, 2), round(T, 2)] for t, T in s["contacts"]])
        else:
            errs = [t for k, t in self._collect_plan_notes() if k == "err"] if P else []
            out["stichplan"] = dict(fehler=errs or ["nicht berechnet"])
        return out

    # ---------- Ereignisse
    def _changed(self, key):
        if self._filling:
            return
        if key not in UI_KEYS:
            self._mark_dirty()
        if key in ("rho", "cp", "lam"):
            m = next((m for m in MATS if m["name"] == self.v["matPreset"].get()), None) if "matPreset" in self.v else None
            if m and (parse(self.v["rho"].get()) != m["rho"] or parse(self.v["cp"].get()) != m["cp"]
                      or parse(self.v["lam"].get()) != m["k"]):
                self._filling = True
                self.v["matPreset"].set(CUSTOM)
                self._filling = False
        if key == "hc" and "furnPreset" in self.v:
            f = next((f for f in FURN if f["name"] == self.v["furnPreset"].get()), None)
            if f and parse(self.v["hc"].get()) != f["hc"]:
                self._filling = True
                self.v["furnPreset"].set(CUSTOM)
                self._filling = False
        self.update_all()

    def _fill_material(self):
        m = next((m for m in MATS if m["name"] == self.v["matPreset"].get()), None)
        if m:
            self._filling = True
            self.v["rho"].set(str(m["rho"]))
            self.v["cp"].set(str(m["cp"]))
            self.v["lam"].set(str(m["k"]))
            self.v["matName"].set(m["name"])
            for k, val in kf_fields(m).items():
                self.v[k].set(val)
            self._filling = False
        self._mark_dirty()
        self.update_all()

    def _fill_furnace(self):
        f = next((f for f in FURN if f["name"] == self.v["furnPreset"].get()), None)
        if f:
            self._filling = True
            self.v["hc"].set(str(f["hc"]))
            self._filling = False
        self._mark_dirty()
        self.update_all()

    def _tab_changed(self):
        idx = self.nb.index(self.nb.select())
        if idx == 1:
            self.run_heat()
        elif idx == 2:
            self.run_plan()
        elif idx == 3:
            self.run_sim_preview()
        self._schedule_save()

    def _schedule_save(self):
        if self._after_save:
            self.after_cancel(self._after_save)
        self._after_save = self.after(800, self._save)

    # ---------- Werte lesen
    def g(self, key):
        return parse(self.v[key].get())

    def read_mat(self):
        rho, cp, k = self.g("rho"), self.g("cp"), self.g("lam")
        bad = [self.mark(self.e_rho, not rho > 0), self.mark(self.e_cp, not cp > 0), self.mark(self.e_lam, not k > 0)]
        name = self.v["matName"].get().strip() or "Werkstoff"
        return dict(rho=rho, cp=cp, k=k, name=name, ok=not any(bad))

    def read_block(self):
        shape = self.v["blkShape"].get()
        d, L = self.g("blkD"), self.g("blkL")
        r = 0.0 if shape == "rund" else (0.0 if not self.v["blkR"].get().strip() else self.g("blkR"))
        ok = True
        ok = not self.mark(self.e_blkD, not d > 0) and ok
        ok = not self.mark(self.e_blkL, not L > 0) and ok
        badr = shape != "rund" and not (r >= 0 and (not d > 0 or r <= d / 2))
        ok = not self.mark(self.e_blkR, badr) and ok
        return dict(shape=shape, d=d, L=L, r=r, ok=ok)

    # ---------- Schmieden
    def compute_forge(self):
        mat, b = self.read_mat(), self.read_block()
        shape, mode = self.v["tgtShape"].get(), self.v["tgtMode"].get()
        loss = 0.0 if not self.v["loss"].get().strip() else self.g("loss")
        r = 0.0 if shape == "rund" else (0.0 if not self.v["tgtR"].get().strip() else self.g("tgtR"))
        errs, warns = [], []
        if self.mark(self.e_loss, not (0 <= loss < 100)):
            errs.append("Materialverlust muss zwischen 0 und unter 100 % liegen.")
        if not b["ok"]:
            errs.append("Rohblock: Maße prüfen. Der Kantenradius darf höchstens die halbe Kantenlänge "
                        "bzw. Schlüsselweite betragen.")
        d1 = L1 = A1 = None
        if mode == "dim":
            d1 = self.g("tgtD")
            self.mark(self.e_tgtL, False)
            if self.mark(self.e_tgtD, not d1 > 0):
                errs.append("Querschnittsmaß des Schmiedestücks eingeben.")
        else:
            L1 = self.g("tgtL")
            self.mark(self.e_tgtD, False)
            if self.mark(self.e_tgtL, not L1 > 0):
                errs.append("Endlänge des Schmiedestücks eingeben.")
        if self.mark(self.e_tgtR, shape != "rund" and not r >= 0):
            errs.append("Kantenradius des Schmiedestücks prüfen.")
        if errs:
            return dict(ok=False, errs=errs, b=b, mat=mat)
        A0 = area(b["shape"], b["d"], b["r"])
        V0 = A0 * b["L"]
        Veff = V0 * (1 - loss / 100.0)
        if mode == "dim":
            if shape != "rund" and r > d1 / 2:
                self.mark(self.e_tgtR, True)
                return dict(ok=False, errs=["Der Kantenradius des Schmiedestücks ist größer als das halbe "
                                            "Querschnittsmaß."], b=b, mat=mat)
            A1 = area(shape, d1, r)
            L1 = Veff / A1
        else:
            A1 = Veff / L1
            d1 = dim_from_area(shape, A1, r)
            if shape != "rund" and r > d1 / 2:
                self.mark(self.e_tgtR, True)
                return dict(ok=False, errs=["Bei dieser Endlänge ist der Kantenradius größer als das halbe "
                                            "Querschnittsmaß."], b=b, mat=mat)
        ratio = A0 / A1
        phi = math.log(ratio)
        if ratio < 1 - 1e-9:
            sl = b["L"] / b["d"]
            txt = f"Der Querschnitt wird größer. Das ist ein Stauchvorgang mit φ = {nf(phi, 2)}."
            if sl > 2.5:
                txt += f" Schlankheit h₀/d₀ = {nf(sl, 2)} liegt über 2,5: Knickgefahr beim Stauchen (Faustregel)."
            warns.append(("warn", txt))
        elif abs(ratio - 1) < 1e-9:
            warns.append(("info", "Ausgangs- und Endquerschnitt sind gleich groß, es findet keine Streckung statt."))
        return dict(ok=True, b=b, mat=mat, shape=shape, mode=mode, d1=d1, r1=r, L1=L1, A0=A0, A1=A1,
                    V0=V0, Veff=Veff, loss=loss, ratio=ratio, phi=phi, warns=warns)

    def _rows(self, parent, rows, start=0):
        sc = self.sc
        for i, (lab, val, sub) in enumerate(rows):
            r = start + i
            tk.Label(parent, text=lab, bg=C["paper"], fg=C["muted"], font=self.f_base, anchor="w").grid(
                row=r, column=0, sticky="w", pady=(int(3 * sc), 0))
            tk.Label(parent, text=val, bg=C["paper"], fg=C["ink"], font=self.f_val, anchor="e").grid(
                row=r, column=1, sticky="e", pady=(int(3 * sc), 0))
            if sub:
                tk.Label(parent, text=sub, bg=C["paper"], fg=C["muted"], font=self.f_small, anchor="e").grid(
                    row=r, column=2, sticky="e", padx=(int(10 * sc), 0), pady=(int(3 * sc), 0))
        parent.columnconfigure(0, weight=1)

    def _tile_wrap(self, width):
        # Umbruchbreite aus der Breite des Containers (nicht der Kachel selbst) -> keine Rückkopplung
        wl = max(80, int(width / 4 - 40 * self.sc))
        for lab in self._tile_labels:
            if lab.winfo_exists() and int(lab.cget("wraplength")) != wl:
                lab.configure(wraplength=wl)

    def _tiles(self, parent, rows):
        sc = self.sc
        self._tile_labels = []
        for i, (lab, val, sub) in enumerate(rows):
            r, c = i // 4, i % 4
            tile = tk.Frame(parent, bg=C["line"])
            tile.grid(row=r, column=c, sticky="nsew", padx=(0, int(10 * sc)), pady=(0, int(8 * sc)))
            inner = tk.Frame(tile, bg=C["paper"], padx=int(8 * sc))
            inner.pack(fill="both", expand=True, padx=(max(1, int(sc)), 0))
            wl = []
            l1 = tk.Label(inner, text=lab, bg=C["paper"], fg=C["muted"], font=self.f_small, anchor="w", justify="left",
                          wraplength=int(120 * sc))
            l1.pack(anchor="w")
            wl.append(l1)
            tk.Label(inner, text=val, bg=C["paper"], fg=C["ink"], font=self.f_sub, anchor="w").pack(anchor="w")
            if sub:
                l2 = tk.Label(inner, text=sub, bg=C["paper"], fg=C["muted"], font=self.f_small, anchor="w",
                              justify="left", wraplength=int(120 * sc))
                l2.pack(anchor="w")
                wl.append(l2)
            self._tile_labels.extend(wl)
        for c in range(4):
            parent.columnconfigure(c, weight=1, uniform="tile")
        if parent.winfo_width() > 1:
            self._tile_wrap(parent.winfo_width())

    def _rows2(self, parent, rows):
        sc = self.sc
        tk.Frame(parent, bg=C["line"], height=1).grid(row=0, column=0, columnspan=7, sticky="ew", pady=(0, int(2 * sc)))
        for i, (lab, val, sub) in enumerate(rows):
            r, c = 1 + i // 2, (i % 2) * 4
            tk.Label(parent, text=lab, bg=C["paper"], fg=C["muted"], font=self.f_base, anchor="w").grid(
                row=r, column=c, sticky="w", pady=(int(2 * sc), 0))
            tk.Label(parent, text=val, bg=C["paper"], fg=C["ink"], font=self.f_val, anchor="e").grid(
                row=r, column=c + 1, sticky="e", padx=(int(10 * sc), 0), pady=(int(2 * sc), 0))
            tk.Label(parent, text=sub, bg=C["paper"], fg=C["muted"], font=self.f_small, anchor="w").grid(
                row=r, column=c + 2, sticky="w", padx=(int(8 * sc), 0), pady=(int(2 * sc), 0))
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(4, weight=1)
        parent.columnconfigure(3, minsize=int(36 * sc))

    def _note(self, parent, kind, text, row):
        bgc = {"warn": C["warn_bg"], "err": C["err_bg"], "info": C["info_bg"]}[kind]
        bar = {"warn": C["warn"], "err": C["err"], "info": C["cold"]}[kind]
        fr = tk.Frame(parent, bg=bar)
        fr.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(int(8 * self.sc), 0))
        tk.Label(fr, text=text, bg=bgc, fg=C["ink"], font=self.f_base, justify="left", anchor="w",
                 wraplength=int(520 * self.sc), padx=int(10 * self.sc), pady=int(6 * self.sc)).pack(
            fill="x", padx=(int(3 * self.sc), 0))

    def render_block_info(self, f):
        for w in self.blk_info.winfo_children():
            w.destroy()
        b = f["b"]
        rows = []
        if b["ok"]:
            A0 = area(b["shape"], b["d"], b["r"])
            V = A0 * b["L"]
            rows += [("Querschnitt A₀", nf(A0 / 100, 1) + " cm²"), ("Volumen", nf(V / 1e6, 2) + " dm³")]
            if f["mat"]["ok"]:
                rows.append(("Masse", nf(f["mat"]["rho"] * V * 1e-9, 1) + " kg"))
            if b["shape"] == "acht":
                rows += [("Eckmaß", nf(b["d"] / math.cos(math.pi / 8), 1) + " mm"),
                         ("Seitenlänge", nf(b["d"] * (SQ2 - 1), 1) + " mm")]
        for i, (a, v) in enumerate(rows):
            tk.Label(self.blk_info, text=a, bg=C["panel"], fg=C["muted"], font=self.f_base).grid(row=i, column=0, sticky="w")
            tk.Label(self.blk_info, text=v, bg=C["panel"], fg=C["ink"], font=self.f_val).grid(row=i, column=1, sticky="e")
        self.blk_info.columnconfigure(0, weight=1)

    def render_forge(self, f):
        p = self.res_f
        for w in p.winfo_children():
            w.destroy()
        if not f["ok"]:
            for i, e in enumerate(f["errs"]):
                self._note(p, "err", e, i)
            return
        if f["mode"] == "dim":
            what, big = "Endlänge L₁", nf(f["L1"], 0) + " mm"
        else:
            what = {"rund": "Erforderlicher Durchmesser D₁", "vier": "Erforderliche Kantenlänge a₁",
                    "acht": "Erforderliche Schlüsselweite SW₁"}[f["shape"]]
            big = nf(f["d1"], 1) + " mm"
        tk.Label(p, text=what, bg=C["paper"], fg=C["muted"], font=self.f_sub).grid(row=0, column=0, columnspan=3, sticky="w")
        tk.Label(p, text=big, bg=C["paper"], fg=C["cold"], font=self.f_big).grid(row=1, column=0, columnspan=3, sticky="w")
        line = geo_label(f["shape"], f["d1"], f["L1"]) + (f", r = {nf(f['r1'], 1)} mm" if f["r1"] > 0 else "")
        tk.Label(p, text=line, bg=C["paper"], font=self.f_base).grid(row=2, column=0, columnspan=3, sticky="w",
                                                                       pady=(0, int(6 * self.sc)))
        dL = f["L1"] - f["b"]["L"]
        rows = [
            ("Querschnitt Rohblock A₀", nf(f["A0"] / 100, 1) + " cm²", ""),
            ("Querschnitt Schmiedestück A₁", nf(f["A1"] / 100, 1) + " cm²", ""),
            ("Verschmiedungsgrad A₀/A₁", nf(f["ratio"], 2) + " : 1", ""),
            ("Umformgrad φ = ln(A₀/A₁)", nf(f["phi"], 3), ""),
            ("Längenänderung", ("+" if dL >= 0 else "−") + nf(abs(dL), 0) + " mm", ""),
            ("Volumen Schmiedestück", nf(f["Veff"] / 1e6, 2) + " dm³",
             f"nach {nf(f['loss'], 1)} % Verlust" if f["loss"] > 0 else ""),
        ]
        if f["mat"]["ok"]:
            rows.append(("Masse Schmiedestück", nf(f["mat"]["rho"] * f["Veff"] * 1e-9, 1) + " kg", ""))
        if f["shape"] == "acht":
            rows.append(("Eckmaß / Seitenlänge",
                         nf(f["d1"] / math.cos(math.pi / 8), 1) + " / " + nf(f["d1"] * (SQ2 - 1), 1) + " mm", ""))
        self._rows(p, rows, start=3)
        for i, (kind, txt) in enumerate(f["warns"]):
            self._note(p, kind, txt, 3 + len(rows) + i)

    # ---------- Zeichnung
    def _section_poly(self, shape, d, r, cx, cy, s):
        R = d / 2 * s
        rr = (r or 0) * s
        if shape == "vier":
            V = [(cx - R, cy - R), (cx + R, cy - R), (cx + R, cy + R), (cx - R, cy + R)]
        else:
            Rc = R / math.cos(math.pi / 8)
            V = [(cx + Rc * math.cos(math.pi / 8 + i * math.pi / 4), cy + Rc * math.sin(math.pi / 8 + i * math.pi / 4))
                 for i in range(8)]
        if rr <= 0.3:
            return [c for p in V for c in p]
        tl = rr if shape == "vier" else rr * math.tan(math.pi / 8)
        pts = []
        n = len(V)
        for i in range(n):
            P, A, B = V[i], V[i - 1], V[(i + 1) % n]
            la = math.hypot(A[0] - P[0], A[1] - P[1])
            lb = math.hypot(B[0] - P[0], B[1] - P[1])
            pin = (P[0] + (A[0] - P[0]) * tl / la, P[1] + (A[1] - P[1]) * tl / la)
            pout = (P[0] + (B[0] - P[0]) * tl / lb, P[1] + (B[1] - P[1]) * tl / lb)
            # Mittelpunkt der Rundung: von pin senkrecht zur Kante ins Innere
            ux, uy = (P[0] - A[0]) / la, (P[1] - A[1]) / la
            nx, ny = -uy, ux
            if nx * (cx - pin[0]) + ny * (cy - pin[1]) < 0:
                nx, ny = -nx, -ny
            mcx, mcy = pin[0] + nx * rr, pin[1] + ny * rr
            a0 = math.atan2(pin[1] - mcy, pin[0] - mcx)
            a1 = math.atan2(pout[1] - mcy, pout[0] - mcx)
            da = (a1 - a0 + math.pi) % (2 * math.pi) - math.pi
            for j in range(9):
                a = a0 + da * j / 8
                pts += [mcx + rr * math.cos(a), mcy + rr * math.sin(a)]
        return pts

    def _shape_on(self, cv, shape, d, r, cx, cy, s, **kw):
        if shape == "rund":
            R = d / 2 * s
            return cv.create_oval(cx - R, cy - R, cx + R, cy + R, **kw)
        return cv.create_polygon(self._section_poly(shape, d, r, cx, cy, s), **kw)

    def _side_bar(self, cv, shape, d, r, L, x, yc, s):
        w = max(L * s, 2)
        hh = max(d * s, 2) / 2
        y0, y1 = yc - hh, yc + hh
        if shape == "rund":
            stops = [(0, C["al0"]), (0.28, C["al1"]), (0.62, C["al2"]), (1, C["al0"])]
            n = int(max(2, min(80, 2 * hh)))
            for i in range(n):
                tpos = (i + 0.5) / n
                for j in range(len(stops) - 1):
                    if stops[j][0] <= tpos <= stops[j + 1][0]:
                        u = (tpos - stops[j][0]) / (stops[j + 1][0] - stops[j][0])
                        col = mix(stops[j][1], stops[j + 1][1], u)
                        break
                cv.create_rectangle(x, y0 + 2 * hh * i / n, x + w, y0 + 2 * hh * (i + 1) / n + 1, fill=col, outline="")
        elif shape == "vier":
            cv.create_rectangle(x, y0, x + w, y1, fill=C["al_face"], outline="")
            if r > 0 and r * s > 1.5:
                e = hh - r * s
                cv.create_line(x, yc - e, x + w, yc - e, fill=C["al2"])
                cv.create_line(x, yc + e, x + w, yc + e, fill=C["al2"])
        else:
            half = d * (SQ2 - 1) / 2 * s
            cv.create_rectangle(x, y0, x + w, y1, fill=C["al_flat"], outline="")
            cv.create_rectangle(x, yc - half, x + w, yc + half, fill=C["al_face"], outline="")
            cv.create_line(x, yc - half, x + w, yc - half, fill=C["al0"])
            cv.create_line(x, yc + half, x + w, yc + half, fill=C["al0"])
        cv.create_rectangle(x, y0, x + w, y1, outline=C["ink"], width=1.2)
        cv.create_line(x - 12, yc, x + w + 12, yc, fill=C["muted"], dash=(12, 3, 2, 3))
        return y1, w

    def _dim_line(self, cv, x0, x1, yref, ydim, text, strong):
        col = C["cold"] if strong else C["ink"]
        font = self.f_cv_b if strong else self.f_cv
        cv.create_line(x0, yref + 3, x0, ydim + 5, fill=C["muted"])
        cv.create_line(x1, yref + 3, x1, ydim + 5, fill=C["muted"])
        sc = self.sc
        cv.create_line(x0, ydim, x1, ydim, fill=col, arrow=tk.BOTH if x1 - x0 > 20 else tk.NONE,
                       arrowshape=(int(9 * sc), int(9 * sc), int(3.5 * sc)))
        tw = font.measure(text)
        if x1 - x0 > tw + 24:
            cv.create_text((x0 + x1) / 2, ydim - 3, text=text, anchor="s", fill=col, font=font)
        else:
            cv.create_text(x1 + 8, ydim, text=text, anchor="w", fill=col, font=font)

    def draw_forge(self):
        cv = self.cv_f
        cv.delete("all")
        f = self.forge
        if not f or not f["ok"]:
            return
        sc = self.sc
        W, H = max(cv.winfo_width(), 300), max(cv.winfo_height(), 200)
        b = f["b"]
        sec = min(H - 80 * sc, 240 * sc, W * 0.3)
        sec = max(sec, 80 * sc)
        sq = (sec - 16 * sc) / max(b["d"], f["d1"])
        scx, scy = 16 * sc + sec / 2, 26 * sc + sec / 2
        cv.create_text(16 * sc, 4 * sc, text="Querschnitte", anchor="nw", font=self.f_cv_b, fill=C["ink"])
        cv.create_line(scx - sec / 2 - 4, scy, scx + sec / 2 + 4, scy, fill=C["muted"], dash=(12, 3, 2, 3))
        cv.create_line(scx, scy - sec / 2 - 4, scx, scy + sec / 2 + 4, fill=C["muted"], dash=(12, 3, 2, 3))
        self._shape_on(cv, f["shape"], f["d1"], f["r1"], scx, scy, sq, fill=C["cold_soft"], outline=C["cold"], width=1.6)
        self._shape_on(cv, b["shape"], b["d"], b["r"], scx, scy, sq, fill="", outline=C["ink"], width=1.2, dash=(5, 3))
        ly = scy + sec / 2 + 18 * sc
        lx = scx - sec / 2
        cv.create_line(lx, ly, lx + 20 * sc, ly, fill=C["ink"], width=1.2, dash=(5, 3))
        cv.create_text(lx + 28 * sc, ly, text=f"Rohblock {SHAPES[b['shape']]['pre']} {fdim(b['d'])}", anchor="w",
                       font=self.f_cv, fill=C["ink"])
        cv.create_rectangle(lx, ly + 13 * sc, lx + 20 * sc, ly + 23 * sc, fill=C["cold_soft"], outline=C["cold"], width=1.4)
        cv.create_text(lx + 28 * sc, ly + 18 * sc, text=f"Schmiedestück {SHAPES[f['shape']]['pre']} {fdim(f['d1'])}",
                       anchor="w", font=self.f_cv, fill=C["ink"])

        bx = sec + 70 * sc
        bw = W - bx - 30 * sc
        if bw < 60:
            return
        # Höhe begrenzt den Maßstab ebenfalls
        avail_h = H - 2 * (26 + 42) * sc - 10 * sc
        sl = min(bw / max(b["L"], f["L1"]), max(avail_h, 20) / (b["d"] + f["d1"]))
        h0, h1 = max(b["d"] * sl, 2), max(f["d1"] * sl, 2)
        yc0 = 26 * sc + h0 / 2
        cv.create_text(bx, yc0 - h0 / 2 - 4 * sc, text="Rohblock", anchor="sw", font=self.f_cv_b, fill=C["ink"])
        y1, w0 = self._side_bar(cv, b["shape"], b["d"], b["r"], b["L"], bx, yc0, sl)
        self._dim_line(cv, bx, bx + w0, y1, y1 + 22 * sc, nf(b["L"], 0) + " mm", False)
        yc1 = y1 + 42 * sc + 26 * sc + h1 / 2
        cv.create_text(bx, yc1 - h1 / 2 - 4 * sc, text="Schmiedestück", anchor="sw", font=self.f_cv_b, fill=C["ink"])
        y2, w1 = self._side_bar(cv, f["shape"], f["d1"], f["r1"], f["L1"], bx, yc1, sl)
        self._dim_line(cv, bx, bx + w1, y2, y2 + 22 * sc, nf(f["L1"], 0) + " mm", True)

    # ---------- Aufheizen
    def schedule_heat(self):
        if self._after_heat:
            self.after_cancel(self._after_heat)
        self._after_heat = self.after(180, self.run_heat)

    def run_heat(self):
        self._after_heat = None
        p = self.res_h
        for w in list(p.winfo_children()) + list(self.res_h2.winfo_children()):
            w.destroy()
        f = self.forge
        mat = self.read_mat()
        T0, Tf, Tz, hc, eps = self.g("T0"), self.g("Tf"), self.g("Tz"), self.g("hc"), self.g("eps")
        errs = []
        bad = [self.mark(self.e_T0, not math.isfinite(T0)), self.mark(self.e_Tf, not math.isfinite(Tf)),
               self.mark(self.e_Tz, not math.isfinite(Tz))]
        if any(bad):
            errs.append("Temperaturen vollständig eingeben.")
        elif Tf <= T0:
            self.mark(self.e_Tf, True)
            errs.append("Die Ofentemperatur muss über der Einsatztemperatur liegen.")
        elif Tz >= Tf:
            self.mark(self.e_Tz, True)
            errs.append("Die Zieltemperatur muss unter der Ofentemperatur liegen, sonst wird sie rechnerisch nie "
                        "erreicht. Üblich ist ein Abstand von 5 bis 20 K.")
        elif Tz <= T0:
            self.mark(self.e_Tz, True)
            errs.append("Die Zieltemperatur liegt nicht über der Einsatztemperatur.")
        if self.mark(self.e_hc, not hc >= 0):
            errs.append("Wärmeübergangskoeffizient prüfen.")
        if self.mark(self.e_eps, not (0 <= eps <= 1)):
            errs.append("Der Emissionsgrad muss zwischen 0 und 1 liegen.")
        if hc == 0 and eps == 0:
            errs.append("Ohne Konvektion und Strahlung findet kein Wärmeübergang statt.")
        if not mat["ok"]:
            errs.append("Stoffwerte des Werkstoffs prüfen.")
        geo = None
        if self.v["heatGeo"].get() == "blk":
            if f["b"]["ok"]:
                b = f["b"]
                geo = dict(shape=b["shape"], d=b["d"], r=b["r"], L=b["L"], label="Rohblock")
            else:
                errs.append("Maße des Rohblocks prüfen.")
        else:
            if f["ok"]:
                geo = dict(shape=f["shape"], d=f["d1"], r=f["r1"], L=f["L1"], label="Schmiedestück")
            else:
                errs.append("Die Geometrie des Schmiedestücks ist noch nicht vollständig berechnet.")
        self.lb_geo.configure(text=("Gerechnet wird mit " + geo_label(geo["shape"], geo["d"], geo["L"]) + ".") if geo else "")
        fh = next((x for x in FURN if x["name"] == self.v["furnPreset"].get()), None)
        self.lb_furn.configure(text=fh["hint"] if fh else "Eigene Werte für den Wärmeübergang.")
        if errs:
            for i, e in enumerate(errs):
                self._note(p, "err", e, i)
            self.heat = None
            self.lb_chart.configure(text="Temperaturverlauf")
            self.draw_chart()
            return

        res = simulate_heating(geo["shape"], geo["d"] / 1000, geo["L"] / 1000, mat["rho"], mat["cp"], mat["k"],
                               T0, Tf, Tz, hc, eps)
        mass = mat["rho"] * area(geo["shape"], geo["d"], geo["r"]) * geo["L"] * 1e-9
        tk.Label(p, text=f"Aufheizzeit bis Kern {nf(Tz, 0)} °C", bg=C["paper"], fg=C["muted"], font=self.f_sub).grid(
            row=0, column=0, columnspan=3, sticky="w")
        if res["t_reach"] is None:
            tk.Label(p, text="> 100 h", bg=C["paper"], fg=C["heat"], font=self.f_big).grid(row=1, column=0, sticky="w")
            self._note(p, "warn", "Die Zieltemperatur wird innerhalb von 100 h nicht erreicht. Wärmeübergang und "
                                  "Stoffwerte prüfen.", 2)
            summary = None
        else:
            Q = mass * mat["cp"] * (res["T_mean"] - T0) / 3.6e6
            tk.Label(p, text=fmt_time(res["t_reach"]), bg=C["paper"], fg=C["heat"], font=self.f_big).grid(
                row=1, column=0, columnspan=3, sticky="w")
            tk.Label(p, text=f"{mat['name']}, {geo['label']}\n{geo_label(geo['shape'], geo['d'], geo['L'])}",
                     bg=C["paper"], font=self.f_base, justify="left", wraplength=int(290 * self.sc)).grid(
                row=2, column=0, columnspan=3, sticky="w")
            rows = [
                (f"Oberfläche erreicht {nf(Tz, 0)} °C", fmt_time(res["t_surf"]), "an der kältesten Stelle"),
                ("Größte ΔT im Block", nf(res["max_dt"], 1) + " K", "nach " + fmt_time(res["t_max_dt"])),
                ("Biot-Zahl (mit V/A)", nf(res["Bi"], 3),
                 "gleichmäßige Durchwärmung" if res["Bi"] < 0.1 else "deutliche Gradienten"),
                ("Wärmeübergang h bei Ofentemp.", nf(res["h_max"], 1) + " W/(m² K)",
                 "davon Strahlung " + nf(res["h_max"] - hc, 1)),
                ("Zeitkonstante τ = ρ·c·V/(h·A)", fmt_time(res["tau"]), ""),
                ("Spezifische Aufheizzeit", nf(res["t_reach"] / 60 / geo["d"], 2) + " min/mm", ""),
                ("Masse", nf(mass, 1) + " kg", ""),
                ("Aufgenommene Wärmemenge", nf(Q, 2) + " kWh", "bis Kern = Ziel"),
            ]
            self._tiles(self.res_h2, rows)
        self.heat = dict(res=res, T0=T0, Tf=Tf, Tz=Tz, geo=geo, mat=mat, hc=hc, eps=eps, mass=mass)
        self.lg_labels[2].configure(text=("Ecke" if geo["shape"] == "vier" else "Stirnkante") + ", wärmste Stelle")
        self.lb_chart.configure(text=f"Temperaturverlauf {mat['name']}, {geo['label']} "
                                     f"{geo_label(geo['shape'], geo['d'], geo['L'])}")
        self.draw_chart()
        self.sf_h.refit()

    # ---------- Diagramm
    def draw_chart(self):
        cv = self.cv_h
        cv.delete("all")
        self._chart = None
        if not self.heat:
            return
        sc = self.sc
        res, T0, Tf, Tz = self.heat["res"], self.heat["T0"], self.heat["Tf"], self.heat["Tz"]
        W, H = max(cv.winfo_width(), 300), max(cv.winfo_height(), 160)
        ml, mr, mt, mb = 58 * sc, 98 * sc, 14 * sc, 40 * sc
        pw, ph = W - ml - mr, H - mt - mb
        if pw < 50 or ph < 40:
            return
        t_end = res["t"][-1] or 1.0
        use_h = t_end > 3 * 3600
        tu = 3600.0 if use_h else 60.0
        xmax = t_end / tu
        dT = Tf - T0
        ymin, ymax = T0 - 0.04 * dT, Tf + 0.06 * dT

        def X(t):
            return ml + (t / tu) / xmax * pw

        def Y(T):
            return mt + (ymax - T) / (ymax - ymin) * ph

        ys = nice_step(ymax - ymin, max(3, int(ph / (45 * sc))))
        v = math.ceil(ymin / ys) * ys
        while v <= ymax + 1e-9:
            cv.create_line(ml, Y(v), ml + pw, Y(v), fill=C["line"])
            cv.create_text(ml - 8, Y(v), text=nf(v, 0), anchor="e", fill=C["muted"], font=self.f_small)
            v += ys
        xs = nice_step(xmax, max(3, int(pw / (90 * sc))))
        v = 0.0
        while v <= xmax + 1e-9:
            cv.create_line(X(v * tu), mt, X(v * tu), mt + ph, fill=C["line"])
            cv.create_text(X(v * tu), mt + ph + 6, text=nf(v, 1 if xs < 1 else 0), anchor="n", fill=C["muted"],
                           font=self.f_small)
            v += xs
        cv.create_text(ml + pw / 2, H - 4, text="Zeit in " + ("h" if use_h else "min"), anchor="s", fill=C["muted"],
                       font=self.f_small)
        cv.create_text(14 * sc, mt + ph / 2, text="Temperatur in °C", angle=90, fill=C["muted"], font=self.f_small)
        cv.create_line(ml, Y(Tf), ml + pw, Y(Tf), fill=C["ink"], dash=(7, 4), width=1.2)
        cv.create_line(ml, Y(Tz), ml + pw, Y(Tz), fill=C["heat"], dash=(2, 3))
        cv.create_text(ml + pw + 6, Y(Tf) - 2, text=f"Ofen {nf(Tf, 0)} °C", anchor="sw", fill=C["ink"], font=self.f_small)
        cv.create_text(ml + pw + 6, Y(Tz) + 2, text=f"Ziel {nf(Tz, 0)} °C", anchor="nw", fill=C["heat"], font=self.f_small)

        n = len(res["t"])
        step = max(1, int(n / (pw * 1.5)))
        idx = list(range(0, n, step))
        if idx[-1] != n - 1:
            idx.append(n - 1)

        def pts(arr):
            out = []
            for i in idx:
                out += [X(res["t"][i]), Y(arr[i])]
            return out

        cv.create_line(pts(res["edge"]), fill=C["heat3"], width=1.5, dash=(2, 3))
        cv.create_line(pts(res["surf"]), fill=C["heat2"], width=2)
        cv.create_line(pts(res["core"]), fill=C["heat"], width=2.6)
        if res["t_reach"] is not None:
            xr = X(res["t_reach"])
            cv.create_line(xr, Y(Tz), xr, mt + ph, fill=C["heat"])
            cv.create_oval(xr - 4, Y(Tz) - 4, xr + 4, Y(Tz) + 4, fill=C["heat"], outline="")
            cv.create_text(xr - 7, mt + ph - 8, text=fmt_time(res["t_reach"]), anchor="se", fill=C["heat"],
                           font=self.f_cv_b)
        self._chart = dict(X=X, Y=Y, ml=ml, pw=pw, mt=mt, ph=ph, tu=tu, xmax=xmax)

    def _hover(self, ev):
        ch = self._chart
        if not ch or not self.heat:
            return
        res = self.heat["res"]
        if not (ch["ml"] <= ev.x <= ch["ml"] + ch["pw"]):
            self._hover_clear()
            return
        tt = (ev.x - ch["ml"]) / ch["pw"] * ch["xmax"] * ch["tu"]
        ts = res["t"]
        lo, hi = 0, len(ts) - 1
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if ts[mid] < tt:
                lo = mid
            else:
                hi = mid
        i = lo if tt - ts[lo] < ts[hi] - tt else hi
        x = ch["X"](ts[i])
        cv = self.cv_h
        cv.delete("hover")
        cv.create_line(x, ch["mt"], x, ch["mt"] + ch["ph"], fill=C["muted"], tags="hover")
        for arr, col in ((res["core"], C["heat"]), (res["surf"], C["heat2"]), (res["edge"], C["heat3"])):
            y = ch["Y"](arr[i])
            cv.create_oval(x - 3.5, y - 3.5, x + 3.5, y + 3.5, fill=col, outline="", tags="hover")
        edge = "Ecke" if self.heat["geo"]["shape"] == "vier" else "Stirnkante"
        self.lb_read.configure(fg=C["ink"], text=f"t = {fmt_time(ts[i])}      Kern {nf(res['core'][i], 1)} °C      "
                                                 f"Oberfläche {nf(res['surf'][i], 1)} °C      {edge} {nf(res['edge'][i], 1)} °C")

    def _hover_clear(self):
        self.cv_h.delete("hover")

    # ---------- Gesamtaktualisierung
    def update_all(self):
        bs, ts, mode = self.v["blkShape"].get(), self.v["tgtShape"].get(), self.v["tgtMode"].get()
        if bs not in SHAPES:
            self.v["blkShape"].set("rund")
            return
        if ts not in SHAPES:
            self.v["tgtShape"].set("vier")
            return
        self.e_blkD._lbl.configure(text=SHAPES[bs]["dim"] + "₀")
        self.e_tgtD._lbl.configure(text=SHAPES[ts]["dim"] + "₁")
        for e, show in ((self.e_blkR, bs != "rund"), (self.e_tgtR, ts != "rund"),
                        (self.e_tgtD, mode == "dim"), (self.e_tgtL, mode == "len")):
            self._show_row(e, show)
        mat = self.read_mat()
        self.lb_alpha.configure(text=("Temperaturleitfähigkeit a = " + nf(mat["k"] / (mat["rho"] * mat["cp"]) * 1e6, 1)
                                      + " mm²/s") if mat["ok"] else "")
        self.forge = self.compute_forge()
        self.render_block_info(self.forge)
        self.render_forge(self.forge)
        self.draw_forge()
        self.sf_f.refit()
        idx = self.nb.index(self.nb.select())
        if idx == 1:
            self.schedule_heat()
        elif idx == 2:
            self.schedule_plan()
        elif idx == 3:
            self.schedule_sim()
        self._schedule_save()

    def _show_row(self, entry, show):
        if not hasattr(entry, "_row_widgets"):
            info = entry.grid_info()
            r = int(info["row"])
            entry._row_widgets = [(w, w.grid_info()) for w in entry.master.grid_slaves(row=r)]
        for w, info in entry._row_widgets:
            if show:
                if not w.winfo_manager():
                    w.grid(**{k: v for k, v in info.items() if k != "in"})
            else:
                w.grid_remove()

    # ================================================================ Stichplan
    PLAN_COLS = (
        # key, header, width (Zeichen), Art
        ("nr", "Nr", 3, "lbl"),
        ("shape", "Form", 8, "in"),
        ("dim", "Maß\n[mm]", 6, "in"),
        ("len", "Länge\n[mm]", 6, "in"),
        ("nrot", "Dreh-\nlagen", 3, "in"),
        ("bk", "b_K\n[mm]", 5, "in"),
        ("phi", "φ", 5, "out"),
        ("strokes", "Hübe", 4, "out"),
        ("tc", "t_K\n[s]", 5, "out"),
        ("ak", "A_K\n[cm²]", 5, "out"),
        ("kf", "k_f\n[MPa]", 5, "out"),
        ("F", "F\n[MN]", 5, "out"),
        ("Tc", "T Kern\n[°C]", 5, "out"),
        ("Ts", "T Ofl.\n[°C]", 5, "out"),
        ("Tk", "T Kont.\n[°C]", 5, "out"),
        ("del", "", 2, "btn"),
    )

    def _flow_layout(self, frame, items):
        """Abschnitte nebeneinander anordnen und bei Platzmangel umbrechen."""
        gap = int(16 * self.sc)

        def relayout(_e=None):
            w = frame.winfo_width()
            if w < 20:
                return
            rows, cur, used = [], [], 0
            for it in items:
                rw = it.winfo_reqwidth() + gap
                if cur and used + rw > w:
                    rows.append(cur)
                    cur, used = [], 0
                cur.append(it)
                used += rw
            if cur:
                rows.append(cur)
            sig = tuple(len(r) for r in rows)
            if getattr(frame, "_sig", None) == sig:
                return
            frame._sig = sig
            for it in items:
                it.grid_forget()
            for r, row in enumerate(rows):
                for c, it in enumerate(row):
                    it.grid(row=r, column=c, sticky="nw", padx=(0, gap), pady=(int(8 * self.sc), 0))

        frame.bind("<Configure>", relayout)
        frame._relayout = relayout

    def _psection(self, parent, title):
        fr = tk.Frame(parent, bg=C["paper"])
        tk.Frame(fr, bg=C["line_strong"], height=1).grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Label(fr, text=title, font=self.f_sub, bg=C["paper"]).grid(row=1, column=0, columnspan=3, sticky="w",
                                                                       pady=(int(4 * self.sc), int(2 * self.sc)))
        fr.columnconfigure(0, weight=1)
        return fr

    def _build_plan_tab(self):
        sc, bg = self.sc, C["paper"]
        t = self.tab_p
        t.columnconfigure(0, weight=1)
        t.rowconfigure(3, weight=1)

        # --- Prozessparameter
        hd = tk.Frame(t, bg=bg)
        hd.grid(row=0, column=0, sticky="ew", pady=(int(10 * sc), 0))
        tk.Frame(hd, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x")
        hrow = tk.Frame(hd, bg=bg)
        hrow.pack(fill="x", pady=(int(6 * sc), 0))
        tk.Label(hrow, text="Prozessparameter", font=self.f_h, bg=bg).pack(side="left")
        self.bt_params = ttk.Button(hrow, text="Ausblenden", command=self._toggle_params)
        self.bt_params.pack(side="right")
        self.lb_kf = tk.Label(hrow, text="", font=self.f_small, fg=C["muted"], bg=bg)
        self.lb_kf.pack(side="right", padx=(0, int(16 * sc)))

        self.par = tk.Frame(t, bg=bg)
        self.par.grid(row=1, column=0, sticky="ew")
        self.pe = {}

        def fld(sec, row, label, key, unit):
            e = self.field(sec, row, label, key, unit, bg, width=7)
            self.pe[key] = e
            return e

        s1 = self._psection(self.par, "Ablauf und Grenzen")
        fld(s1, 2, "T bei Ofenentnahme", "pT0", "°C")
        fld(s1, 3, "Transport zur Presse", "pTrans", "s")
        fld(s1, 4, "Nebenzeit je Hub", "pAux", "s")
        fld(s1, 5, "Pause zw. Stichen", "pPause", "s")
        fld(s1, 6, "T min. Oberfläche", "pTmin", "°C")
        fld(s1, 7, "T max. Kern", "pTmax", "°C")
        self.hint(s1, "Nebenzeit: Lüften, Vorschub, Drehen, Anfahren; Werkstück liegt auf dem Untersattel "
                      "(QForm: CoolingInTool).", bg, wrap=250).grid(
            row=8, column=0, columnspan=3, sticky="w")

        s2 = self._psection(self.par, "Presse und Sättel")
        fld(s2, 2, "Hubgeschwindigkeit v", "pV", "mm/s")
        fld(s2, 3, "Verweilzeit unter Last", "pDwell", "s")
        fld(s2, 4, "Sattelbreite", "pBdie", "mm")
        fld(s2, 5, "Presskraft max.", "pFmax", "MN")
        tk.Label(s2, text="Bisslänge", bg=bg, font=self.f_base).grid(row=6, column=0, sticky="w",
                                                                     pady=(int(4 * sc), 0))
        self.seg(s2, "pBite", [("min", "Sattelbreite"), ("const", "konstant"), ("sbh", "aus s_B/h")], bg).grid(
            row=7, column=0, columnspan=3, sticky="ew", pady=(int(2 * sc), int(2 * sc)))
        fld(s2, 8, "Bisslänge konstant", "pSBc", "mm")
        fld(s2, 9, "s_B / h", "pSB", "–")
        self.hint(s2, "Sattelbreite = ebene Arbeitsfläche ohne Kantenradien. Bisslänge Sattelbreite: minimale "
                      "Bisszahl. Konstant: gleiche Bisslänge für alle Stiche. s_B/h üblich 0,5 bis 0,8. Bisslänge "
                      "höchstens Sattelbreite.", bg, wrap=250).grid(
            row=10, column=0, columnspan=3, sticky="w")

        s3 = self._psection(self.par, "Wärmeübergang")
        fld(s3, 2, "Satteltemperatur", "pTdie", "°C")
        fld(s3, 3, "h_K unter Last", "pHc", "W/(m² K)")
        fld(s3, 4, "h_K beim Aufliegen", "pHrest", "W/(m² K)")
        fld(s3, 5, "λ Sattel", "pLdie", "W/(m K)")
        fld(s3, 6, "Luft h", "pHair", "W/(m² K)")
        fld(s3, 7, "Emissionsgrad ε", "pEps", "–")
        self.hint(s3, "h_K unter Last mit Schmierstoff etwa 10\u00a0000 bis 30\u00a0000 W/(m² K), beim Aufliegen in der "
                      "Nebenzeit deutlich kleiner; λ Warmarbeitsstahl ≈ 30.",
                  bg, wrap=250).grid(row=8, column=0, columnspan=3, sticky="w")

        s4 = self._psection(self.par, "Umformung und Fließspannung")
        self.seg(s4, "kfModel", [("st", "Sellars-Tegart"), ("hs", "Hensel-Spittel")], bg).grid(
            row=2, column=0, columnspan=3, sticky="ew", pady=(0, int(4 * sc)))
        self.kf_st = tk.Frame(s4, bg=bg)
        self.kf_st.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.e_st = {}
        for i, (lab, key, unit) in enumerate((("α", "kfAlpha", "1/MPa"), ("n", "kfN", "–"),
                                              ("Q", "kfQ", "kJ/mol"), ("ln A", "kfLnA", "ln(1/s)"))):
            self.e_st[key] = self.field(self.kf_st, i, lab, key, unit, bg, width=7)
        self.kf_st.columnconfigure(0, weight=1)
        self.kf_hs = tk.Frame(s4, bg=bg)
        self.kf_hs.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.e_hs = {}
        for i, key in enumerate(sp_HS_KEYS):
            r, c = i // 2, (i % 2) * 2
            tk.Label(self.kf_hs, text=("A" if key == "A" else "m" + "₀₁₂₃₄₅₆₇₈₉"[int(key[1:])]), bg=bg,
                     font=self.f_base).grid(row=r, column=c, sticky="w", padx=(0 if c == 0 else int(10 * sc), 4),
                                            pady=int(2 * sc))
            e = ttk.Entry(self.kf_hs, textvariable=self.var("hs" + key), width=10, justify="right")
            e.grid(row=r, column=c + 1, sticky="e", pady=int(2 * sc))
            self.e_hs["hs" + key] = e
        self.hint(self.kf_hs, "T in °C, φ̇ in 1/s, Ergebnis in MPa (Form wie QForm).", bg, wrap=250).grid(
            row=5, column=0, columnspan=4, sticky="w")
        fld(s4, 4, "Wirkungsgrad η_F", "pEta", "–")
        fld(s4, 5, "Reibungszahl μ", "pMu", "–")
        self.hint(s4, "η_F: Formänderungswirkungsgrad (Schiebungen, Reibung), Recken etwa 0,7 bis 0,8.", bg,
                  wrap=250).grid(row=6, column=0, columnspan=3, sticky="w")

        self._plan_sections = [s1, s2, s3, s4]
        self._flow_layout(self.par, self._plan_sections)

        # --- Stichtabelle
        sec, inn = self._section(t, "Stiche")
        sec.grid(row=2, column=0, sticky="ew", pady=(int(14 * sc), 0))
        self.lb_thint = self.hint(
            inn, "Maß (D, a oder SW) oder Länge eintragen, der andere Wert wird berechnet (grau). Drehlagen und "
                 "Kontaktbreite b_K je Sattel werden automatisch belegt und lassen sich überschreiben; Feld leeren "
                 "stellt den Automatikwert wieder her. t_K Kontaktzeit je Hub, A_K Kontaktfläche je Sattel, F Presskraft, "
                 "T Ofl. Oberfläche (Mittel über den Umfang), T Kont. Kontaktzone am Hubende; Temperaturen am "
                 "Stichende. Zeile anklicken für Details.", bg, wrap=1000)
        self.lb_thint.pack(anchor="w")
        self.tbl_scroll = HScroll(inn, bg)
        self.tbl_scroll.pack(fill="x", pady=(int(6 * sc), 0))
        self.tbl = self.tbl_scroll.inner
        brow = tk.Frame(inn, bg=bg)
        brow.pack(fill="x", pady=(int(6 * sc), 0))
        ttk.Button(brow, text="+ Stich anhängen", command=self.add_pass).pack(side="left")
        ttk.Button(brow, text="Als CSV speichern …", command=self.export_plan_csv).pack(side="right")
        ttk.Button(brow, text="Für QForm exportieren …", command=self.export_qform).pack(
            side="right", padx=(0, int(8 * sc)))
        ttk.Button(brow, text="Als Excel-Datei speichern …", command=self.export_plan_xlsx).pack(
            side="right", padx=(0, int(8 * sc)))
        self.lb_detail = tk.Label(inn, text="", font=self.f_base, bg=bg, fg=C["ink"], justify="left", anchor="w",
                                  wraplength=int(1000 * sc))
        self.lb_detail.pack(fill="x", pady=(int(8 * sc), 0))
        self.plan_notes = tk.Frame(inn, bg=bg)
        self.plan_notes.pack(fill="x")
        self.plan_notes.columnconfigure(0, weight=1)
        self._pwrap = int(1000 * sc)

        def rewrap(e):
            wl = max(300, e.width - int(30 * sc))
            if wl != self._pwrap:
                self._pwrap = wl
                for lab in (self.lb_thint, self.lb_detail):
                    lab.configure(wraplength=wl)
                for fr in self.plan_notes.winfo_children():
                    for lab in fr.winfo_children():
                        lab.configure(wraplength=wl)
        inn.bind("<Configure>", rewrap)

        # --- Darstellung
        cf = tk.Frame(t, bg=bg)
        cf.grid(row=3, column=0, sticky="nsew", pady=(int(14 * sc), 0))
        tk.Frame(cf, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x")
        hd2 = tk.Frame(cf, bg=bg)
        hd2.pack(fill="x", pady=(int(6 * sc), int(2 * sc)))
        self.lb_pchart = tk.Label(hd2, text="Temperaturverlauf", font=self.f_h, bg=bg)
        self.lb_pchart.pack(side="left")
        self.seg(hd2, "planView", [("temp", "Temperaturverlauf"), ("geo", "Geometrie")], bg).pack(side="right")
        self.plg = tk.Frame(cf, bg=bg)
        self.plg.pack(fill="x", pady=(0, int(2 * sc)))
        for txt, col, wd, dash, kind in (("Kern", C["heat"], 3, None, "line"),
                                         ("Oberfläche (Mittel über den Umfang)", C["heat2"], 2, None, "line"),
                                         ("Kontaktzone am Hubende", C["heat3"], 0, None, "dot"),
                                         ("Grenzen", C["cold"], 1.5, (6, 3), "line")):
            swc = tk.Canvas(self.plg, width=int(24 * sc), height=int(12 * sc), bg=bg, highlightthickness=0)
            if kind == "dot":
                swc.create_oval(int(9 * sc), int(3 * sc), int(15 * sc), int(9 * sc), fill=col, outline="")
            else:
                kw = dict(dash=dash) if dash else {}
                swc.create_line(0, int(6 * sc), int(24 * sc), int(6 * sc), fill=col, width=wd * sc, **kw)
            swc.pack(side="left", padx=(0, int(5 * sc)))
            tk.Label(self.plg, text=txt, font=self.f_small, fg=C["muted"], bg=bg).pack(side="left",
                                                                                     padx=(0, int(16 * sc)))
        self.cv_p = tk.Canvas(cf, bg=bg, highlightthickness=0, height=int(300 * sc))
        self.lb_pread = tk.Label(cf, text="Mauszeiger über das Diagramm bewegen, um Werte abzulesen.",
                                 font=self.f_base, fg=C["muted"], bg=bg, anchor="w")
        self.lb_pread.pack(side="bottom", fill="x")
        self.cv_p.pack(fill="both", expand=True)
        self.cv_p.bind("<Configure>", lambda e: self.draw_plan())
        self.cv_p.bind("<Motion>", self._phover)
        self.cv_p.bind("<Leave>", lambda e: self.cv_p.delete("hover"))

        hdr = [c[1] for c in self.PLAN_COLS]
        for j, (key, text, w, kind) in enumerate(self.PLAN_COLS):
            tk.Label(self.tbl, text=text, font=self.f_small, fg=C["muted"], bg=bg, justify="center",
                     anchor="s").grid(row=0, column=j, sticky="sew", padx=int(1 * sc), pady=(0, int(2 * sc)))
        tk.Frame(self.tbl, bg=C["line_strong"], height=1).grid(row=1, column=0, columnspan=len(hdr), sticky="ew")

    def _toggle_params(self):
        show = self.v["planParams"].get() != "1"
        self.v["planParams"].set("1" if show else "0")

    def _apply_params_visibility(self):
        show = self.v["planParams"].get() == "1"
        if show and not self.par.winfo_manager():
            self.par.grid(row=1, column=0, sticky="ew")
        elif not show:
            self.par.grid_remove()
        self.bt_params.configure(text="Ausblenden" if show else "Einblenden")
        hs = self.v["kfModel"].get() == "hs"
        if hs:
            self.kf_st.grid_remove()
            self.kf_hs.grid()
        else:
            self.kf_hs.grid_remove()
            self.kf_st.grid()
        if hasattr(self.par, "_relayout"):
            self.par._sig = None
            self.after_idle(self.par._relayout)
        self.sf_p.refit()

    # ---------- Zeilen der Stichtabelle
    def _new_row(self, d):
        row = dict(mode=d.get("mode", "dim"), nrot_man=bool(str(d.get("nrot", "")).strip()),
                   bk_man=bool(str(d.get("bk", "")).strip()), w={})
        self._filling = True
        for key in ("shape", "dim", "len", "nrot", "bk", "rw", "rwT"):
            sv = tk.StringVar(self)
            val = d.get(key, "")
            if key == "rw":
                val = "1" if str(val).strip() in ("1", "true", "True") else "0"
            if key == "shape":
                val = SHAPES.get(val, SHAPES["vier"])["name"]
            sv.set(str(val))
            sv.trace_add("write", lambda *_a, r=row, k=key: self._row_changed(r, k))
            row[key] = sv
        self._filling = False
        return row

    def _set_plan_rows(self, passes):
        self.plan_rows = [self._new_row(p) for p in passes]
        self.plan_sel = min(self.plan_sel, max(0, len(self.plan_rows) - 1))
        self._rebuild_table()

    def _plan_rows_data(self):
        out = []
        for r in self.plan_rows:
            out.append(dict(shape=SHAPE_BY_NAME.get(r["shape"].get(), "vier"), mode=r["mode"],
                            dim=r["dim"].get(), len=r["len"].get(),
                            nrot=r["nrot"].get() if r["nrot_man"] else "",
                            bk=r["bk"].get() if r["bk_man"] else "",
                            rw=r["rw"].get(), rwT=r["rwT"].get()))
        return out

    def _row_changed(self, row, key):
        if self._filling:
            return
        if key in ("rw", "rwT"):                     # Wiedererwärmung: nur Reiter QForm-Simulation betroffen
            self._mark_dirty()
            self.schedule_sim()
            self._schedule_save()
            return
        if key == "dim":
            row["mode"] = "dim"
        elif key == "len":
            row["mode"] = "len"
        elif key == "nrot":
            row["nrot_man"] = bool(row["nrot"].get().strip())
        elif key == "bk":
            row["bk_man"] = bool(row["bk"].get().strip())
        self._mark_dirty()
        self.schedule_plan()
        self._schedule_save()

    def _rebuild_table(self):
        sc, bg = self.sc, C["paper"]
        for w in self.tbl.grid_slaves():
            if int(w.grid_info()["row"]) >= 2:
                w.destroy()
        self.blk_cells = {}
        # Zeile 0: Rohblock
        r0 = 2
        for j, (key, _t, width, kind) in enumerate(self.PLAN_COLS):
            if kind == "btn":
                continue
            lab = tk.Label(self.tbl, text="", font=self.f_base, bg=bg, fg=C["muted"], width=width,
                           anchor="e" if kind != "lbl" else "center")
            lab.grid(row=r0, column=j, sticky="ew", padx=int(1 * sc), pady=int(2 * sc))
            self.blk_cells[key] = lab
        self.blk_cells["nr"].configure(text="0")
        for i, row in enumerate(self.plan_rows):
            rr = 3 + i
            row["w"] = {}
            for j, (key, _t, width, kind) in enumerate(self.PLAN_COLS):
                if key == "nr":
                    w = tk.Label(self.tbl, text=str(i + 1), font=self.f_val, bg=bg, width=width, cursor="hand2")
                elif key == "shape":
                    w = ttk.Combobox(self.tbl, textvariable=row["shape"], state="readonly", width=width,
                                     values=[S["name"] for S in SHAPES.values()])
                elif kind == "in":
                    w = ttk.Entry(self.tbl, textvariable=row[key], width=width, justify="right")
                elif kind == "out":
                    w = tk.Label(self.tbl, text="–", font=self.f_base, bg=bg, width=width, anchor="e", cursor="hand2")
                else:
                    w = ttk.Button(self.tbl, text="✕", width=2, command=lambda r=row: self.del_pass(r))
                w.grid(row=rr, column=j, sticky="ew", padx=int(1 * sc), pady=int(1 * sc))
                if kind in ("out", "lbl"):
                    w.bind("<Button-1>", lambda e, k=i: self._select_pass(k))
                elif kind == "in":
                    w.bind("<FocusIn>", lambda e, k=i: self._select_pass(k))
                row["w"][key] = w
        self.schedule_plan()

    def _select_pass(self, k, redraw=True):
        if self.plan_sel == k:
            return
        self.plan_sel = k
        self._render_plan_details()
        self._render_plan_notes()
        self._highlight_sel()
        if redraw:
            self.draw_plan()

    def _highlight_sel(self):
        for i, row in enumerate(self.plan_rows):
            bgc = C["cold_soft"] if i == self.plan_sel else C["paper"]
            for key, _t, _w, kind in self.PLAN_COLS:
                if kind in ("out", "lbl") and key in row.get("w", {}):
                    row["w"][key].configure(bg=bgc)

    def add_pass(self):
        last = self.plan_rows[-1] if self.plan_rows else None
        shape = SHAPE_BY_NAME.get(last["shape"].get(), "vier") if last else "vier"
        A_last = None
        if self.plan and self.plan.get("geo"):
            gs = [g for g in self.plan["geo"] if g.get("ok")]
            if gs and len(gs) == len(self.plan_rows):
                A_last = gs[-1]["A"]
        if A_last is None and not self.plan_rows:
            b = self.read_block()
            if b["ok"]:
                A_last = area(b["shape"], b["d"], b["r"])
        if A_last is None:
            d0 = parse(last["dim"].get()) if last else float("nan")
            A_last = sp.area(shape, d0) if d0 and d0 > 0 else sp.area(shape, 100.0)
        # Vorschlag: Verschmiedungsgrad etwa 1,3 : 1 in der Form des letzten Stichs, auf 5 mm gerundet
        dim = max(5.0, round(sp.dim_from_area(shape, A_last / 1.3) / 5.0) * 5.0)
        self.plan_rows.append(self._new_row(dict(shape=shape, mode="dim", dim=fe(dim, 0))))
        self.plan_sel = len(self.plan_rows) - 1
        self._mark_dirty()
        self._rebuild_table()
        self._schedule_save()

    def del_pass(self, row):
        if row in self.plan_rows:
            self.plan_rows.remove(row)
            self.plan_sel = min(self.plan_sel, max(0, len(self.plan_rows) - 1))
            self._mark_dirty()
            self._rebuild_table()
            self._schedule_save()

    # ---------- Rechnung
    def schedule_plan(self):
        if self._after_plan:
            self.after_cancel(self._after_plan)
        self._after_plan = self.after(250, self.run_plan)

    def read_kf(self):
        if self.v["kfModel"].get() == "hs":
            p, ok = dict(type="hs"), True
            for key, e in self.e_hs.items():
                val = self.g(key)
                good = math.isfinite(val) and (val > 0 if key == "hsA" else True)
                self.mark(e, not good)
                ok = ok and good
                p[key[2:]] = val
            return p, ok
        p = dict(type="st", alpha=self.g("kfAlpha"), n=self.g("kfN"), Q=self.g("kfQ"), lnA=self.g("kfLnA"))
        ok = True
        for key, e in self.e_st.items():
            val = self.g(key)
            good = math.isfinite(val) and (val > 0 if key != "kfLnA" else True)
            self.mark(e, not good)
            ok = ok and good
        return p, ok

    def read_proc(self):
        spec = dict(pT0=("T_start", lambda x: True), pTrans=("t_transfer", lambda x: x >= 0),
                    pAux=("t_aux", lambda x: x >= 0), pPause=("t_pause", lambda x: x >= 0),
                    pV=("v", lambda x: x > 0), pDwell=("t_dwell", lambda x: x >= 0),
                    pBdie=("b_die", lambda x: x > 0),
                    pTdie=("T_die", lambda x: True), pHc=("h_c", lambda x: x >= 0),
                    pHrest=("h_rest", lambda x: x >= 0),
                    pLdie=("lam_die", lambda x: x > 0), pHair=("h_air", lambda x: x >= 0),
                    pEps=("eps", lambda x: 0 <= x <= 1), pEta=("eta_F", lambda x: 0 < x <= 1),
                    pMu=("mu", lambda x: x >= 0), pFmax=("F_max", lambda x: x > 0),
                    pTmin=("T_min", lambda x: True), pTmax=("T_max", lambda x: True))
        proc, ok = {}, True
        for key, (name, cond) in spec.items():
            val = self.g(key)
            good = math.isfinite(val) and cond(val)
            self.mark(self.pe[key], not good)
            ok = ok and good
            proc[name] = val
        mode = self.v["pBite"].get()
        mode = mode if mode in ("min", "const", "sbh") else "min"
        proc["bite_mode"] = mode
        # nur das zum Bisslängenmodus gehörende Feld wird geprüft; das andere erscheint grau
        for key, name, need, cond in (("pSBc", "sB_const", "const", lambda x: x > 0),
                                      ("pSB", "sB_rel", "sbh", lambda x: 0 < x <= 3)):
            val = self.g(key)
            good = math.isfinite(val) and cond(val)
            if mode == need:
                self.mark(self.pe[key], not good)
                ok = ok and good
            else:
                self.pe[key].configure(style="Calc.TEntry")
            proc[name] = val if good else None
        return proc, ok

    def run_plan(self):
        self._after_plan = None
        self._apply_params_visibility()
        mat = self.read_mat()
        b = self.read_block()
        kf, kf_ok = self.read_kf()
        proc, proc_ok = self.read_proc()
        errs = []
        if not mat["ok"]:
            errs.append("Stoffwerte des Werkstoffs prüfen (linke Spalte).")
        if not b["ok"]:
            errs.append("Maße des Rohblocks prüfen (linke Spalte).")
        if not kf_ok:
            errs.append("Parameter der Fließspannung prüfen.")
        if not proc_ok:
            errs.append("Prozessparameter prüfen (rot markierte Felder).")
        if kf_ok and math.isfinite(proc.get("T_start", float("nan"))):
            try:
                kv = sp.kf_mean(kf, proc["T_start"], 0.5, 0.5, 1.0)
                self.lb_kf.configure(text=f"Kontrolle: k_f({nf(proc['T_start'], 0)} °C; φ = 0,5; φ̇ = 1 s⁻¹) = "
                                          f"{nf(kv, 1)} MPa")
            except (ValueError, OverflowError, ZeroDivisionError):
                self.lb_kf.configure(text="")
        passes = []
        for r in self.plan_rows:
            nrot = parse(r["nrot"].get()) if r["nrot_man"] else None
            bk = parse(r["bk"].get()) if r["bk_man"] else None
            passes.append(dict(shape=SHAPE_BY_NAME.get(r["shape"].get(), "vier"), mode=r["mode"],
                               dim=parse(r["dim"].get()), len=parse(r["len"].get()),
                               nrot=int(round(nrot)) if nrot and math.isfinite(nrot) and nrot >= 1 else None,
                               bk=bk if bk and math.isfinite(bk) and bk > 0 else None,
                               nrot_bad=r["nrot_man"] and not (nrot and math.isfinite(nrot) and nrot >= 1),
                               bk_bad=r["bk_man"] and not (bk and math.isfinite(bk) and bk > 0)))
        geo, res = [], None
        if b["ok"]:
            A0 = area(b["shape"], b["d"], b["r"])
            block = dict(shape=b["shape"], d=b["d"], L=b["L"], A=A0)
            geo = sp.plan_geometry(block, [dict(p, dim=p["dim"] if math.isfinite(p["dim"]) else None,
                                                len=p["len"] if math.isfinite(p["len"]) else None)
                                           for p in passes])
            for g, p in zip(geo, passes):
                g["nrot"], g["bk"] = p["nrot"], p["bk"]
            if not errs and geo and any(g.get("ok") for g in geo):
                try:
                    res = sp.simulate_plan(block, geo, mat, kf, proc)
                except (ValueError, OverflowError, ZeroDivisionError) as ex:
                    errs.append(f"Die Temperaturrechnung ist fehlgeschlagen ({ex}). Eingaben prüfen.")
        self.plan = dict(geo=geo, res=res, proc=proc if proc_ok else None, mat=mat, block=b, errs=errs,
                         passes=passes, kf=kf)
        self._render_plan_table()
        self._render_plan_details()
        self._render_plan_notes()
        self.sf_p.refit()
        self.lb_pchart.configure(text=f"Temperaturverlauf {mat['name']}, Stichplan"
                                 if self.v["planView"].get() == "temp" else f"Geometrie {mat['name']}, Stichplan")
        if self.v["planView"].get() != "temp":
            self.plg.pack_forget()
            self.lb_pread.pack_forget()
        else:
            self.plg.pack(fill="x", pady=(0, int(2 * self.sc)), before=self.cv_p)
            self.lb_pread.pack(side="bottom", fill="x", before=self.cv_p)
        self.draw_plan()

    def _set_calc(self, entry, sv, text, calc):
        """Feldinhalt programmgesteuert setzen (ohne Moduswechsel) und Darstellung als berechnet/eingegeben."""
        if calc and sv.get() != text:
            self._filling = True
            sv.set(text)
            self._filling = False
        entry.configure(style="Calc.TEntry" if calc else "TEntry")

    def _render_plan_table(self):
        P = self.plan
        b = P["block"]
        res = P["res"]
        proc = P["proc"]
        cells = self.blk_cells
        if b["ok"]:
            cells["shape"].configure(text=SHAPES[b["shape"]]["name"])
            cells["dim"].configure(text=fdim(b["d"]))
            cells["len"].configure(text=nf(b["L"], 0))
        for key in ("phi", "strokes", "tc", "ak", "kf", "F", "nrot", "bk"):
            cells[key].configure(text="")
        if res:
            cells["Tc"].configure(text=nf(res["start"][0], 0))
            cells["Ts"].configure(text=nf(res["start"][1], 0))
            cells["Tk"].configure(text="")
        else:
            for key in ("Tc", "Ts", "Tk"):
                cells[key].configure(text="")
        geo = P["geo"]
        rows_res = res["rows"] if res else []
        for i, row in enumerate(self.plan_rows):
            w = row["w"]
            g = geo[i] if i < len(geo) else None
            p = P["passes"][i]
            shape = SHAPE_BY_NAME.get(row["shape"].get(), "vier")
            # Maß / Länge
            if g and "d" in g:
                if row["mode"] == "dim":
                    self._set_calc(w["len"], row["len"], fe(g["L"], 0), True)
                    w["dim"].configure(style="TEntry")
                else:
                    self._set_calc(w["dim"], row["dim"], fe(g["d"], 1), True)
                    w["len"].configure(style="TEntry")
            else:
                bad_in = "dim" if row["mode"] == "dim" else "len"
                calc = "len" if bad_in == "dim" else "dim"
                w[bad_in].configure(style="Bad.TEntry" if g is not None and not g.get("ok") and "d" not in g
                                    else "TEntry")
                self._set_calc(w[calc], row[calc], "", True)
            if g and g.get("ok") is False and "d" in g:
                w["dim" if row["mode"] == "dim" else "len"].configure(style="Bad.TEntry")
            # Drehlagen / Kontaktbreite
            if row["nrot_man"]:
                w["nrot"].configure(style="Bad.TEntry" if p["nrot_bad"] else "TEntry")
            else:
                self._set_calc(w["nrot"], row["nrot"], str(sp.DEFAULT_NROT[shape]), True)
            r = rows_res[i] if i < len(rows_res) else None
            if row["bk_man"]:
                w["bk"].configure(style="Bad.TEntry" if p["bk_bad"] else "TEntry")
            elif r:
                self._set_calc(w["bk"], row["bk"], fe(r["bc_auto"], 0), True)
            else:
                self._set_calc(w["bk"], row["bk"], "", True)
            # Ergebnisse
            vals = dict(phi="–", strokes="–", tc="–", ak="–", kf="–", F="–", Tc="–", Ts="–", Tk="–")
            cols = dict.fromkeys(vals, C["ink"])
            fonts = dict.fromkeys(vals, self.f_base)
            if g and g.get("ok"):
                vals["phi"] = nf(g["phi"], 3)
            if r:
                vals.update(strokes=str(r["strokes"]), tc=nf(r["t_c"], 2), ak=nf(r["A_K"] / 100, 0),
                            kf=nf(r["kf"], 1), F=nf(r["F"], 2), Tc=nf(r["T_core"], 0), Ts=nf(r["T_surf"], 0),
                            Tk=nf(r["T_contact"], 0))
                if proc:
                    if r["F"] > proc["F_max"]:
                        cols["F"], fonts["F"] = C["err"], self.f_val
                    if r["T_core"] > proc["T_max"]:
                        cols["Tc"], fonts["Tc"] = C["err"], self.f_val
                    if r["T_surf"] < proc["T_min"]:
                        cols["Ts"], fonts["Ts"] = C["warn"], self.f_val
                    if r["T_contact"] < proc["T_min"]:
                        cols["Tk"], fonts["Tk"] = C["warn"], self.f_val
            for key, val in vals.items():
                w[key].configure(text=val, fg=cols[key], font=fonts[key])
        self._highlight_sel()

    def _render_plan_details(self):
        P = self.plan
        if not P:
            return
        res = P["res"]
        i = self.plan_sel
        geo = P["geo"]
        if not res or i >= len(res["rows"]):
            g = geo[i] if i < len(geo) else None
            if g and g.get("msg"):
                self.lb_detail.configure(text=f"Stich {i + 1}: {g['msg']}", fg=C["err"])
            else:
                self.lb_detail.configure(text="", fg=C["ink"])
            return
        r, g = res["rows"][i], geo[i]
        dT_air = r["dT_mean"] - r["dT_def"] + r["dT_die"]
        txt = (f"Stich {i + 1}: {SHAPES[g['prev']['shape']]['pre']} {fdim(g['prev']['d'])} → "
               f"{SHAPES[g['shape']]['pre']} {fdim(g['d'])} mm, Länge {nf(g['prev']['L'], 0)} → {nf(g['L'], 0)} mm, "
               f"Verschmiedungsgrad {nf(g['ratio'], 2)} : 1.   "
               f"{r['n_b']} Bisse × {r['n_rot']} Drehlagen ({nf(r['rot_step'], 0)}°) = {r['strokes']} Hübe, "
               f"Bisslänge s_B = {nf(r['sB'], 0)} mm (s_B/h = {nf(r['sB'] / g['h_prev'], 2)}"
               f"{', letzter Biss ' + nf(r['s_last'], 0) + ' mm' if abs(r['s_last'] - r['sB']) > 0.05 else ''}), "
               f"Vorschub {nf(r['feed'], 1)} mm, Hubweg {nf(r['dh'], 0)} mm, Kontaktzeit {nf(r['t_c'], 2)} s, "
               f"φ̇ ≈ {nf(r['phid'], 2)} s⁻¹, Stichdauer {fmt_time(r['t_pass'])}.   "
               f"Mitteltemperatur {('+' if r['dT_mean'] >= 0 else '−')}{nf(abs(r['dT_mean']), 1)} K: "
               f"Umformwärme +{nf(r['dT_def'], 1)} K, Sattelkontakt −{nf(r['dT_die'], 1)} K, "
               f"Luft −{nf(max(0.0, -dT_air), 1)} K.")
        self.lb_detail.configure(text=txt, fg=C["ink"])

    def _render_plan_notes(self):
        p = self.plan_notes
        for w in p.winfo_children():
            w.destroy()
        notes = self._collect_plan_notes(skip_sel=True)
        k = 0
        for kind in ("err", "warn", "info"):
            texts = [t for kk, t in notes if kk == kind]
            if not texts:
                continue
            self._note(p, kind, "\n".join(texts), k)
            lab = p.grid_slaves(row=k)[0].winfo_children()[0]
            lab.configure(wraplength=self._pwrap, pady=int(4 * self.sc))
            k += 1

    def _collect_plan_notes(self, skip_sel=False):
        P = self.plan
        notes = [("err", e) for e in P["errs"]]
        geo, res, proc = P["geo"], P["res"], P["proc"]
        for i, g in enumerate(geo):
            if g.get("msg") and not (skip_sel and i == self.plan_sel):
                notes.append(("err", f"Stich {i + 1}: {g['msg']}"))
        if res and proc:
            for i, r in enumerate(res["rows"]):
                g = geo[i]
                if r["F"] > proc["F_max"]:
                    notes.append(("err", f"Stich {i + 1}: geschätzte Presskraft {nf(r['F'], 2)} MN über der Grenze "
                                         f"von {nf(proc['F_max'], 1)} MN. Kleinere Bisslänge oder Kontaktbreite wählen."))
                if r["T_core"] > proc["T_max"]:
                    notes.append(("err", f"Stich {i + 1}: Kerntemperatur {nf(r['T_core'], 0)} °C über "
                                         f"{nf(proc['T_max'], 0)} °C."))
                if r["T_surf"] < proc["T_min"]:
                    notes.append(("warn", f"Stich {i + 1}: mittlere Oberflächentemperatur {nf(r['T_surf'], 0)} °C unter "
                                          f"{nf(proc['T_min'], 0)} °C. Zwischenwärmen vorsehen."))
                elif r["T_contact"] < proc["T_min"]:
                    notes.append(("warn", f"Stich {i + 1}: Kontaktzone kühlt am Hubende auf {nf(r['T_contact'], 0)} °C "
                                          f"(Grenze {nf(proc['T_min'], 0)} °C)."))
                rel = r["sB"] / g["h_prev"]
                if rel < 0.5 - 1e-6:
                    notes.append(("warn", f"Stich {i + 1}: s_B/h = {nf(rel, 2)} liegt unter 0,5. "
                                          "Gefahr von Überlappungen und schwacher Kerndurchschmiedung."))
            if proc.get("bite_mode") == "const" and proc.get("sB_const") and proc["sB_const"] > proc["b_die"]:
                notes.append(("warn", f"Die konstante Bisslänge {nf(proc['sB_const'], 0)} mm ist größer als die "
                                      f"Sattelbreite {nf(proc['b_die'], 0)} mm; gerechnet wird mit "
                                      f"{nf(proc['b_die'], 0)} mm."))
            ok_geo = [g for g in geo if g.get("ok")]
            if ok_geo and res["rows"]:
                last, rl = ok_geo[-1], res["rows"][len(ok_geo) - 1]
                total = math.log(P["geo"][0]["prev"]["A"] / last["A"])
                notes.append(("info", f"Ergebnis: {geo_label(last['shape'], last['d'], last['L'])}, Verschmiedungsgrad "
                                      f"{nf(math.exp(total), 2)} : 1 (φ = {nf(total, 2)}), "
                                      f"{sum(r['strokes'] for r in res['rows'])} Hübe, Dauer ab Ofenentnahme "
                                      f"{fmt_time(rl['t_end'])}. Am Ende: Kern {nf(rl['T_core'], 0)} °C, "
                                      f"Oberfläche {nf(rl['T_surf'], 0)} °C."))
        return notes

    # ---------- Darstellung
    def draw_plan(self):
        cv = self.cv_p
        cv.delete("all")
        self._pchart = None
        P = self.plan
        if not P:
            return
        if self.v["planView"].get() == "geo":
            self._draw_plan_geo()
        else:
            self._draw_plan_temp()

    def _draw_plan_temp(self):
        cv, sc = self.cv_p, self.sc
        P = self.plan
        res, proc = P["res"], P["proc"]
        if not res or len(res["t"]) < 2:
            return
        W, H = max(cv.winfo_width(), 300), max(cv.winfo_height(), 180)
        ml, mr, mt, mb = 58 * sc, 110 * sc, 22 * sc, 40 * sc
        pw, ph = W - ml - mr, H - mt - mb
        if pw < 60 or ph < 60:
            return
        t_end = res["t"][-1] or 1.0
        tu = 60.0
        xmax = t_end / tu
        vals = res["core"] + res["surf"] + [c[1] for c in res["contacts"]]
        lo, hi = min(vals), max(vals)
        if proc:
            lo = min(lo, proc["T_min"]) if proc["T_min"] > lo - 60 else lo
            hi = max(hi, proc["T_max"]) if proc["T_max"] < hi + 40 else hi
        span = max(hi - lo, 20.0)
        ymin, ymax = lo - 0.06 * span, hi + 0.08 * span

        def X(tt):
            return ml + (tt / tu) / xmax * pw

        def Y(T):
            return mt + (ymax - T) / (ymax - ymin) * ph

        for k, (t0, t1) in enumerate(res["bands"]):
            cv.create_rectangle(X(t0), mt, X(t1), mt + ph, fill=C["panel"] if k % 2 == 0 else C["info_bg"],
                                outline="")
            if X(t1) - X(t0) > 18 * sc:
                cv.create_text((X(t0) + X(t1)) / 2, mt - 4, text=f"S{k + 1}", anchor="s", fill=C["muted"],
                               font=self.f_cv_b)
        ys = nice_step(ymax - ymin, max(3, int(ph / (45 * sc))))
        v = math.ceil(ymin / ys) * ys
        while v <= ymax + 1e-9:
            cv.create_line(ml, Y(v), ml + pw, Y(v), fill=C["line"])
            cv.create_text(ml - 8, Y(v), text=nf(v, 0), anchor="e", fill=C["muted"], font=self.f_small)
            v += ys
        xs = nice_step(xmax, max(3, int(pw / (90 * sc))))
        v = 0.0
        while v <= xmax + 1e-9:
            cv.create_line(X(v * tu), mt + ph, X(v * tu), mt + ph + 4, fill=C["line_strong"])
            cv.create_text(X(v * tu), mt + ph + 6, text=nf(v, 1 if xs < 1 else 0), anchor="n", fill=C["muted"],
                           font=self.f_small)
            v += xs
        cv.create_line(ml, mt + ph, ml + pw, mt + ph, fill=C["line_strong"])
        cv.create_text(ml + pw / 2, H - 4, text="Zeit ab Ofenentnahme in min", anchor="s", fill=C["muted"],
                       font=self.f_small)
        cv.create_text(14 * sc, mt + ph / 2, text="Temperatur in °C", angle=90, fill=C["muted"], font=self.f_small)
        if proc:
            for T, txt in ((proc["T_min"], f"min. {nf(proc['T_min'], 0)} °C"),
                           (proc["T_max"], f"max. {nf(proc['T_max'], 0)} °C")):
                if ymin <= T <= ymax:
                    cv.create_line(ml, Y(T), ml + pw, Y(T), fill=C["cold"], dash=(6, 3), width=1.5)
                    cv.create_text(ml + pw + 6, Y(T), text=txt, anchor="w", fill=C["cold"], font=self.f_small)
        n = len(res["t"])
        step = max(1, int(n / (pw * 1.5)))
        idx = list(range(0, n, step))
        if idx[-1] != n - 1:
            idx.append(n - 1)

        def pts(arr):
            out = []
            for i in idx:
                out += [X(res["t"][i]), Y(arr[i])]
            return out

        cv.create_line(pts(res["surf"]), fill=C["heat2"], width=2)
        cv.create_line(pts(res["core"]), fill=C["heat"], width=2.6)
        rr = 2.6 * sc
        for tt, Tc in res["contacts"]:
            cv.create_oval(X(tt) - rr, Y(Tc) - rr, X(tt) + rr, Y(Tc) + rr, fill=C["heat3"], outline="")
        # Auswahl hervorheben
        if self.plan_sel < len(res["bands"]):
            t0, t1 = res["bands"][self.plan_sel]
            cv.create_rectangle(X(t0), mt, X(t1), mt + ph, outline=C["cold"], width=1.5)
        self._pchart = dict(X=X, Y=Y, ml=ml, pw=pw, mt=mt, ph=ph, tu=tu, xmax=xmax)

    def _phover(self, ev):
        ch = self._pchart
        P = self.plan
        if not ch or not P or not P["res"]:
            return
        res = P["res"]
        if not (ch["ml"] <= ev.x <= ch["ml"] + ch["pw"]):
            self.cv_p.delete("hover")
            return
        tt = (ev.x - ch["ml"]) / ch["pw"] * ch["xmax"] * ch["tu"]
        ts = res["t"]
        lo, hi = 0, len(ts) - 1
        while hi - lo > 1:
            mid = (lo + hi) // 2
            if ts[mid] < tt:
                lo = mid
            else:
                hi = mid
        i = lo if tt - ts[lo] < ts[hi] - tt else hi
        x = ch["X"](ts[i])
        cv = self.cv_p
        cv.delete("hover")
        cv.create_line(x, ch["mt"], x, ch["mt"] + ch["ph"], fill=C["muted"], tags="hover")
        for arr, col in ((res["core"], C["heat"]), (res["surf"], C["heat2"])):
            y = ch["Y"](arr[i])
            cv.create_oval(x - 3.5, y - 3.5, x + 3.5, y + 3.5, fill=col, outline="", tags="hover")
        where = "Transport"
        for k, (t0, t1) in enumerate(res["bands"]):
            if t0 <= ts[i] <= t1:
                where = f"Stich {k + 1}"
                break
            if ts[i] > t1:
                where = f"Pause nach Stich {k + 1}"
        self.lb_pread.configure(fg=C["ink"], text=f"t = {nf(ts[i] / 60, 2)} min ({where})      Kern "
                                                  f"{nf(res['core'][i], 1)} °C      Oberfläche {nf(res['surf'][i], 1)} °C")

    def _draw_plan_geo(self):
        cv, sc = self.cv_p, self.sc
        P = self.plan
        b = P["block"]
        if not b["ok"]:
            return
        stages = [dict(shape=b["shape"], d=b["d"], r=b["r"], L=b["L"], name="Rohblock")]
        for i, g in enumerate(P["geo"]):
            if g.get("ok"):
                stages.append(dict(shape=g["shape"], d=g["d"], r=0.0, L=g["L"], name=f"Stich {i + 1}"))
        W, H = max(cv.winfo_width(), 300), max(cv.winfo_height(), 180)
        lab_w = 78 * sc
        gap = 14 * sc
        maxd = max(s["d"] for s in stages)
        maxL = max(s["L"] for s in stages)
        avail_h = H - 12 * sc - gap * (len(stages) - 1)
        sumd = sum(max(s["d"], 0.35 * maxd) for s in stages)
        txt_w = 150 * sc
        s_len = (W - lab_w - 2 * gap - txt_w) / (maxL + maxd)
        s_h = avail_h / sumd
        s = max(1e-3, min(s_len, s_h))
        sec_w = maxd * s
        y = 6 * sc
        for k, st in enumerate(stages):
            hrow = max(st["d"], 0.35 * maxd) * s
            yc = y + hrow / 2
            sel = (k - 1) == self.plan_sel
            cv.create_text(0, yc, text=st["name"], anchor="w", font=self.f_cv_b if sel else self.f_cv,
                           fill=C["cold"] if sel else C["ink"])
            cx = lab_w + sec_w / 2
            kw = dict(fill=C["cold_soft"], outline=C["cold"], width=1.4) if sel else \
                dict(fill=C["al_face"], outline=C["ink"], width=1.0)
            self._shape_on(cv, st["shape"], st["d"], st["r"], cx, yc, s, **kw)
            bx = lab_w + sec_w + gap
            self._side_bar(cv, st["shape"], st["d"], st["r"], st["L"], bx, yc, s)
            cv.create_text(bx + st["L"] * s + 18 * sc, yc,
                           text=f"{SHAPES[st['shape']]['pre']} {fdim(st['d'])} × {nf(st['L'], 0)} mm",
                           anchor="w", font=self.f_cv_b if sel else self.f_cv, fill=C["cold"] if sel else C["ink"])
            y += hrow + gap

    # ---------- Export Stichplan
    def _plan_export_data(self):
        """Alle Angaben des gerechneten Stichplans als einfache Datenstruktur für den Excel-Export."""
        P = self.plan
        res, geo, b, mat, proc = P["res"], P["geo"], P["block"], P["mat"], P["proc"]
        A0 = area(b["shape"], b["d"], b["r"])
        rows = []
        for i, r in enumerate(res["rows"]):
            g, pr = geo[i], self.plan_rows[i]
            rows.append(dict(r, nr=i + 1, shape=g["shape"], mode=g["mode"], d=g["d"], L=g["L"], A=g["A"],
                             ratio=g["ratio"], phi=g["phi"], phi_cum=g["phi_cum"], h_prev=g["h_prev"],
                             nrot_manual=pr["nrot_man"], bk_manual=pr["bk_man"]))
        model = []
        for block in MODEL_TEXT.split("\n\n"):
            head, _, text = block.partition("\n")
            if head.startswith("Stichplan"):
                model.append((head.strip(), text.strip()))
        return dict(
            version=VERSION, created=datetime.datetime.now(),
            mat=dict(name=mat["name"], rho=mat["rho"], cp=mat["cp"], k=mat["k"]),
            block=dict(shape=b["shape"], d=b["d"], r=b["r"], L=b["L"], A=A0, mass=mat["rho"] * A0 * b["L"] * 1e-9),
            project=self._project_name() if self.project_path else None,
            proc=dict(proc), kf=dict(P["kf"]), start=res["start"], t_transfer=proc["t_transfer"], rows=rows,
            series=dict(t=res["t"], core=res["core"], surf=res["surf"], contacts=res["contacts"]),
            notes=self._collect_plan_notes(), model=model)

    def export_plan_xlsx(self):
        P = self.plan
        if not P or not P["res"] or not P["res"]["rows"]:
            messagebox.showinfo("Excel-Export", "Es liegt noch kein gerechneter Stichplan vor.", parent=self)
            return
        try:
            import xlsx_export
        except ImportError:
            messagebox.showerror("Excel-Export", "Für den Excel-Export wird das Paket openpyxl benötigt "
                                                 "(pip install openpyxl).", parent=self)
            return
        name = "".join(c if c.isalnum() or c in "-_" else "_" for c in P["mat"]["name"])
        init = (f"{self._project_name()}_Stichplan.xlsx" if self.project_path
                else f"Stichplan_{name}_{datetime.date.today():%Y-%m-%d}.xlsx")
        path = filedialog.asksaveasfilename(parent=self, title="Stichplan als Excel-Datei speichern",
                                            defaultextension=".xlsx", initialfile=init,
                                            initialdir=self._project_dir(),
                                            filetypes=[("Excel-Arbeitsmappe", "*.xlsx"), ("Alle Dateien", "*.*")])
        if not path:
            return
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            xlsx_export.export_stichplan(path, self._plan_export_data())
        except PermissionError:
            messagebox.showerror("Excel-Export", "Die Datei kann nicht geschrieben werden. Ist sie noch in Excel "
                                                 "geöffnet? Bitte schließen und erneut speichern.", parent=self)
            return
        except OSError as e:
            messagebox.showerror("Excel-Export", f"Die Datei konnte nicht geschrieben werden:\n{e}", parent=self)
            return
        finally:
            self.config(cursor="")
        if sys.platform == "win32":
            if messagebox.askyesno("Excel-Export", f"Stichplan gespeichert:\n{path}\n\nDatei jetzt öffnen?",
                                   parent=self):
                try:
                    os.startfile(path)
                except OSError:
                    pass
        else:
            messagebox.showinfo("Excel-Export", f"Stichplan gespeichert:\n{path}", parent=self)

    # ---------- QForm-Export
    def _qform_options(self):
        g = lambda k, d: self.v[k].get() if k in self.v and self.v[k].get() else d
        return dict(scope=g("qfScope", "stich"), fmt=g("qfFormat", "xlsx"), dec=g("qfDec", "punkt"),
                    sign=-1 if g("qfSign", "neg") == "neg" else 1, ret=g("qfReturn", "1") == "1")

    def _qform_data(self, opt):
        """Hubtabellen je Stich und Angaben für die Einstellungen in QForm."""
        P = self.plan
        res, geo, proc = P["res"], P["geo"], P["proc"]
        rows = res["rows"]
        blows = qform_export.build_blows(rows, proc, sign=opt["sign"], return_to_start=opt["ret"])
        info, first = [], 1
        for i, (r, b) in enumerate(zip(rows, blows)):
            g = geo[i]
            info.append(dict(nr=i + 1, shape=g["shape"], d=g["d"], L=g["L"], blows=len(b), first=first,
                             last=first + len(b) - 1, n_b=r["n_b"], n_rot=r["n_rot"], rot=r["rot_step"],
                             sB=r["sB"], s_last=r["s_last"], feed=opt["sign"] * r["feed"], ret=b[0][2]))
            first += len(b)
        return blows, info

    def export_qform(self):
        P = self.plan
        if not P or not P["res"] or not P["res"]["rows"]:
            messagebox.showinfo("QForm-Export", "Es liegt noch kein gerechneter Stichplan vor.", parent=self)
            return
        if len(P["res"]["rows"]) < len(self.plan_rows):
            messagebox.showinfo("QForm-Export", "Der Stichplan enthält Stiche, die nicht gerechnet werden konnten. "
                                                "Bitte zuerst die markierten Eingaben korrigieren.", parent=self)
            return
        sc, bg = self.sc, C["paper"]
        top = tk.Toplevel(self)
        top.title("Stichplan für QForm exportieren")
        top.configure(bg=bg)
        top.transient(self)
        top.resizable(False, False)
        fr = tk.Frame(top, bg=bg, padx=int(20 * sc), pady=int(16 * sc))
        fr.pack(fill="both", expand=True)
        tk.Label(fr, text="Stichplan für QForm exportieren", font=self.f_h, bg=bg).grid(row=0, column=0, columnspan=2,
                                                                                         sticky="w")
        self.hint(fr, "Hubtabelle für die Registerkarte „Blows“ im Format der QForm-Importvorlage: Blow, "
                      "CoolingInAir, CoolingInTool, Displacement, RotationAngle.", bg, wrap=760).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(int(2 * sc), int(10 * sc)))
        opts = [
            ("Umfang", "qfScope", [("stich", "Je Stich eine Tabelle"), ("plan", "Ganzer Stichplan in einer Tabelle")]),
            ("Dateiformat", "qfFormat", [("xlsx", "Excel (.xlsx)"), ("csv", "CSV (Semikolon)")]),
            ("Dezimalzeichen CSV", "qfDec", [("punkt", "Punkt"), ("komma", "Komma")]),
            ("Vorschubrichtung", "qfSign", [("neg", "negativ, wie Vorlage"), ("pos", "positiv")]),
        ]
        for i, (lab, key, choices) in enumerate(opts):
            if not self.v[key].get():
                self.v[key].set(DEFAULTS[key])
            tk.Label(fr, text=lab, bg=bg, font=self.f_base).grid(row=2 + i, column=0, sticky="w",
                                                                 padx=(0, int(16 * sc)), pady=int(2 * sc))
            self.seg(fr, key, choices, bg).grid(row=2 + i, column=1, sticky="ew", pady=int(2 * sc))
        if not self.v["qfReturn"].get():
            self.v["qfReturn"].set("1")
        ttk.Checkbutton(fr, text="Erster Hub jedes weiteren Stichs fährt an den Stichanfang zurück",
                        variable=self.v["qfReturn"], onvalue="1", offvalue="0").grid(
            row=6, column=1, sticky="w", pady=(int(4 * sc), 0))
        tk.Frame(fr, bg=C["ink"], height=max(2, int(2 * sc))).grid(row=7, column=0, columnspan=2, sticky="ew",
                                                                  pady=(int(12 * sc), int(4 * sc)))
        tbl = tk.Frame(fr, bg=bg)
        tbl.grid(row=8, column=0, columnspan=2, sticky="ew")
        notes = tk.Label(fr, text="", bg=bg, fg=C["ink"], font=self.f_small, justify="left", anchor="w",
                         wraplength=int(760 * sc))
        notes.grid(row=9, column=0, columnspan=2, sticky="w", pady=(int(8 * sc), 0))
        heads = ("Stich", "Schmiedestück", "Endabstand\nSättel [mm]", "Hübe", "Bisse ×\nDrehlagen", "Dreh-\nschritt",
                 "Bisslänge\n[mm]", "Vorschub\n[mm]", "Rückfahrt\n[mm]")

        def refresh(*_a):
            for w in tbl.winfo_children():
                w.destroy()
            opt = self._qform_options()
            _blows, info = self._qform_data(opt)
            for j, h in enumerate(heads):
                tk.Label(tbl, text=h, bg=C["panel"], fg=C["ink"],
                         font=self.f_small, justify="center", padx=int(6 * sc), pady=int(2 * sc)).grid(
                    row=0, column=j, sticky="nsew")
            for i, d in enumerate(info):
                last = f", letzter {nf(d['s_last'], 0)}" if d["n_b"] > 1 and abs(d["s_last"] - d["sB"]) > 0.05 else ""
                vals = (str(d["nr"]), f"{SHAPES[d['shape']]['pre']} {fdim(d['d'])} × {nf(d['L'], 0)}",
                        fdim(d["d"]), f"{d['first']}–{d['last']}" if opt["scope"] == "plan" else str(d["blows"]),
                        f"{d['n_b']} × {d['n_rot']}", f"{nf(d['rot'], 0)}°", nf(d["sB"], 0) + last,
                        nf(d["feed"], 1) if d["n_b"] > 1 else "–", nf(d["ret"], 1) if d["ret"] else "–")
                for j, v in enumerate(vals):
                    tk.Label(tbl, text=v, bg=bg, fg=C["ink"], font=self.f_val if j == 2 else self.f_base,
                             anchor="e" if j not in (1,) else "w", padx=int(6 * sc)).grid(
                        row=i + 1, column=j, sticky="ew", pady=(int(2 * sc), 0))
            proc = self.plan["proc"]
            dist = ("die Stichtabelle für den ganzen Plan braucht in QForm den Abstand „individuell je Hub“"
                    if opt["scope"] == "plan" else "je Stich eine QForm-Operation mit diesem Abstand")
            notes.configure(text=(
                f"CoolingInAir: Transport {nf(proc['t_transfer'], 0)} s vor dem ersten Hub, Pause {nf(proc['t_pause'], 0)} s "
                f"vor dem ersten Hub jedes weiteren Stichs, sonst 0.  CoolingInTool: Nebenzeit "
                f"{nf(proc['t_aux'], 1)} s vor jedem Hub.  In QForm einstellen: Hubgeschwindigkeit "
                f"{nf(proc['v'], 0)} mm/s und Endabstand der Sättel ({dist}). Erster Biss jedes Stichs am "
                f"Stichanfang, Sattel auf der vollen Bisslänge. Die Einstellungen werden zusätzlich als Textdatei "
                f"gespeichert."))

        traces = [(k, self.v[k].trace_add("write", refresh)) for k in ("qfScope", "qfSign", "qfReturn")]
        refresh()

        def close():
            for k, tid in traces:
                try:
                    self.v[k].trace_remove("write", tid)
                except tk.TclError:
                    pass
            top.destroy()

        def do_export():
            opt = self._qform_options()
            if self._write_qform(opt, parent=top):
                close()

        bt = tk.Frame(fr, bg=bg)
        bt.grid(row=10, column=0, columnspan=2, sticky="e", pady=(int(14 * sc), 0))
        ttk.Button(bt, text="Abbrechen", command=close).pack(side="right")
        ttk.Button(bt, text="Exportieren …", command=do_export).pack(side="right", padx=(0, int(8 * sc)))
        top.protocol("WM_DELETE_WINDOW", close)
        top.bind("<Escape>", lambda e: close())
        top.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - top.winfo_reqwidth()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - top.winfo_reqheight()) // 3)
        top.geometry(f"+{x}+{y}")
        try:
            top.grab_set()
        except tk.TclError:
            pass
        return top

    def _write_qform(self, opt, parent=None):
        """Dateien schreiben. Rückgabe True bei Erfolg."""
        parent = parent or self
        ext = ".xlsx" if opt["fmt"] == "xlsx" else ".csv"
        mat = "".join(c if c.isalnum() or c in "-_" else "_" for c in self.plan["mat"]["name"])
        base = self._project_name() if self.project_path else f"Stichplan_{mat}"
        types = [("Excel-Arbeitsmappe", "*.xlsx")] if ext == ".xlsx" else [("CSV", "*.csv")]
        path = filedialog.asksaveasfilename(
            parent=parent, title="QForm-Hubtabelle speichern" + (" (Basisname, je Stich eine Datei)"
                                                                 if opt["scope"] == "stich" else ""),
            defaultextension=ext, initialfile=f"{base}_QForm{ext}", initialdir=self._project_dir(),
            filetypes=types + [("Alle Dateien", "*.*")])
        if not path:
            return False
        stem = path[:-len(ext)] if path.lower().endswith(ext) else path
        blows, info = self._qform_data(opt)
        if opt["scope"] == "stich":
            files = [(f"{stem}_Stich{i + 1}{ext}", b) for i, b in enumerate(blows)]
        else:
            files = [(stem + ext, [x for b in blows for x in b])]
        txt_path = f"{stem}_Einstellungen.txt"
        try:
            for fpath, bl in files:
                if ext == ".xlsx":
                    qform_export.write_xlsx(fpath, bl)
                else:
                    qform_export.write_csv(fpath, bl, decimal="," if opt["dec"] == "komma" else ".")
            with open(txt_path, "w", encoding="utf-8-sig", newline="\r\n" if sys.platform == "win32" else "\n") as fh:
                fh.write(self._qform_settings_text(opt, info, [os.path.basename(f) for f, _b in files]))
        except ImportError:
            messagebox.showerror("QForm-Export", "Für Excel-Dateien wird das Paket openpyxl benötigt "
                                                 "(pip install openpyxl). CSV ist ohne openpyxl möglich.", parent=parent)
            return False
        except PermissionError:
            messagebox.showerror("QForm-Export", "Eine Datei kann nicht geschrieben werden. Ist sie noch in Excel "
                                                 "geöffnet? Bitte schließen und erneut exportieren.", parent=parent)
            return False
        except OSError as e:
            messagebox.showerror("QForm-Export", f"Die Dateien konnten nicht geschrieben werden:\n{e}", parent=parent)
            return False
        names = "\n".join("  " + os.path.basename(f) for f, _b in files)
        messagebox.showinfo("QForm-Export", f"Gespeichert in {os.path.dirname(path)}:\n{names}\n  "
                                            f"{os.path.basename(txt_path)}", parent=parent)
        return True

    def _qform_settings_text(self, opt, info, names):
        P = self.plan
        proc = P["proc"]
        proj = f", Projekt „{self._project_name()}“" if self.project_path else ""
        L = [f"QForm-Hubtabelle: Stichplan {P['mat']['name']}{proj}",
             f"Erstellt am {datetime.datetime.now():%d.%m.%Y, %H:%M} Uhr mit dem Freiformschmieden-Rechner {VERSION}",
             "",
             "Dateien: " + ", ".join(names),
             "Import in QForm: Registerkarte Blows, Tabelle importieren.",
             "",
             "Spalten (QForm wendet alle Werte vor dem jeweiligen Hub an)",
             f"  CoolingInAir [s]   erster Hub: Transport zur Presse ({nf(proc['t_transfer'], 1)} s); erster Hub jedes "
             f"weiteren Stichs: Pause zwischen den Stichen ({nf(proc['t_pause'], 1)} s); sonst 0",
             f"  CoolingInTool [s]  Nebenzeit je Hub ({nf(proc['t_aux'], 1)} s), Werkstück liegt auf dem Untersattel",
             "  Displacement [mm]  Vorschub vor dem ersten Hub jedes neuen Bisses = Bisslänge × A_vorher/A_nachher"
             + (", Vorzeichen negativ wie in der Vorlage" if opt["sign"] < 0 else ", Vorzeichen positiv")
             + ("; im ersten Hub jedes weiteren Stichs Rückfahrt an den Stichanfang" if opt["ret"] else ""),
             "  RotationAngle [°]  Drehschritt 180°/Drehlagen vor jedem Hub außer dem ersten Hub eines Stichs",
             "",
             "In QForm einzustellen",
             f"  Hubgeschwindigkeit der Presse: {nf(proc['v'], 1)} mm/s",
             f"  Sattelbreite: {nf(proc['b_die'], 0)} mm, Satteltemperatur: {nf(proc['T_die'], 0)} °C",
             "  Abbruch über den Abstand der Sättel"
             + (" (individuell je Hub, Werte siehe Hübe unten)" if opt["scope"] == "plan" else " (je Stich eine Operation)"),
             ""]
        for d in info:
            name, pre = SHAPES[d["shape"]]["name"], SHAPES[d["shape"]]["pre"]
            hub = f"Hübe {d['first']}–{d['last']}" if opt["scope"] == "plan" else f"{d['blows']} Hübe"
            line = (f"  Stich {d['nr']}: {name} {pre} {fdim(d['d'])} mm, Länge danach {nf(d['L'], 0)} mm; {hub}; "
                    f"Endabstand der Sättel {fdim(d['d'])} mm; {d['n_b']} Bisse × {d['n_rot']} Drehlagen "
                    f"({nf(d['rot'], 0)}°); Bisslänge {nf(d['sB'], 1)} mm")
            if abs(d["s_last"] - d["sB"]) > 0.05:
                line += f" (letzter Biss {nf(d['s_last'], 1)} mm)"
            if d["n_b"] > 1:
                line += f"; Vorschub {nf(d['feed'], 1)} mm"
            if d["ret"]:
                line += f"; Rückfahrt im ersten Hub {nf(d['ret'], 1)} mm"
            L.append(line)
        L += ["",
              "Ausgangslage: Zu Beginn jedes Stichs steht der Sattel am Anfang des zu reckenden Bereichs und erfasst "
              "die volle Bisslänge. Der Vorschub berücksichtigt die Längung des gerade geschmiedeten Bisses, sodass "
              "jeder Biss auf ungeschmiedeten Werkstoff trifft ("
              + {"min": "Bisslänge = Sattelbreite, minimale Bisszahl",
                 "const": f"konstante Bisslänge {nf(min(proc['sB_const'] or 0, proc['b_die']), 0)} mm",
                 "sbh": f"Bisslänge aus s_B/h = {nf(proc['sB_rel'] or 0, 2)}"}[proc.get("bite_mode", "min")]
              + ")."]
        return "\n".join(L) + "\n"

    def export_plan_csv(self):
        P = self.plan
        if not P or not P["res"] or not P["res"]["rows"]:
            messagebox.showinfo("Export", "Es liegt noch kein gerechneter Stichplan vor.", parent=self)
            return
        name = "".join(c if c.isalnum() or c in "-_" else "_" for c in P["mat"]["name"])
        init = f"{self._project_name()}_Stichplan.csv" if self.project_path else f"Stichplan_{name}.csv"
        path = filedialog.asksaveasfilename(parent=self, title="Stichplan speichern", defaultextension=".csv",
                                            initialfile=init, initialdir=self._project_dir(),
                                            filetypes=[("CSV (Semikolon)", "*.csv"), ("Alle Dateien", "*.*")])
        if not path:
            return

        def de(x, d):
            return f"{x:.{d}f}".replace(".", ",")

        res, geo = P["res"], P["geo"]
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                wr = csv.writer(fh, delimiter=";")
                wr.writerow([f"Stichplan {P['mat']['name']}"])
                wr.writerow(["Stich", "Form", "Maß [mm]", "Länge [mm]", "Querschnitt [cm²]", "Verschmiedungsgrad",
                             "φ Stich", "φ gesamt", "Drehlagen", "Bisse", "Hübe", "Bisslänge [mm]", "Vorschub [mm]",
                             "Kontaktbreite [mm]", "Kontaktfläche je Sattel [cm²]", "Hubweg [mm]", "Kontaktzeit [s]",
                             "φ̇ [1/s]", "k_f [MPa]", "Presskraft [MN]", "Stichdauer [s]", "Zeit Stichende [s]",
                             "ΔT Umformwärme [K]", "ΔT Sattel [K]", "T Kern [°C]", "T Oberfläche Mittel [°C]",
                             "T Kontaktzone min [°C]"])
                b = P["block"]
                wr.writerow([0, SHAPES[b["shape"]]["name"], de(b["d"], 1), de(b["L"], 0),
                             de(geo[0]["prev"]["A"] / 100, 1)] + [""] * 19 +
                            [de(res["start"][0], 1), de(res["start"][1], 1), ""])
                for i, r in enumerate(res["rows"]):
                    g = geo[i]
                    wr.writerow([i + 1, SHAPES[g["shape"]]["name"], de(g["d"], 1), de(g["L"], 0), de(g["A"] / 100, 1),
                                 de(g["ratio"], 3), de(g["phi"], 4), de(g["phi_cum"], 4), r["n_rot"], r["n_b"],
                                 r["strokes"], de(r["sB"], 1), de(r["feed"], 1), de(r["bc"], 1), de(r["A_K"] / 100, 1),
                                 de(r["dh"], 1),
                                 de(r["t_c"], 3), de(r["phid"], 4), de(r["kf"], 2), de(r["F"], 3), de(r["t_pass"], 1),
                                 de(r["t_end"], 1), de(r["dT_def"], 2), de(-r["dT_die"], 2), de(r["T_core"], 1),
                                 de(r["T_surf"], 1), de(r["T_contact"], 1)])
                wr.writerow([])
                wr.writerow(["Zeit [s]", "T Kern [°C]", "T Oberfläche Mittel [°C]"])
                n = len(res["t"])
                stride = max(1, n // 2000)
                for i in list(range(0, n, stride)) + ([n - 1] if (n - 1) % stride else []):
                    wr.writerow([de(res["t"][i], 2), de(res["core"][i], 2), de(res["surf"][i], 2)])
        except OSError as e:
            messagebox.showerror("Export", f"Die Datei konnte nicht geschrieben werden:\n{e}", parent=self)

    # ---------- Export
    def export_csv(self):
        if not self.heat:
            messagebox.showinfo("Export", "Es liegt noch kein gültiger Temperaturverlauf vor. "
                                          "Bitte zuerst die Aufheizrechnung durchführen.", parent=self)
            return
        h = self.heat
        res = h["res"]
        name = "".join(c if c.isalnum() or c in "-_" else "_" for c in h["mat"]["name"])
        path = filedialog.asksaveasfilename(parent=self, title="Aufheizkurve speichern", defaultextension=".csv",
                                            initialfile=f"Aufheizkurve_{name}.csv",
                                            filetypes=[("CSV (Semikolon)", "*.csv"), ("Alle Dateien", "*.*")])
        if not path:
            return

        def de(x, d):
            return f"{x:.{d}f}".replace(".", ",")

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                wr = csv.writer(fh, delimiter=";")
                wr.writerow(["Zeit [s]", "Zeit [min]", "T Kern [°C]", "T Oberfläche kälteste Stelle [°C]",
                             "T Kante/Ecke [°C]"])
                for i in range(len(res["t"])):
                    wr.writerow([de(res["t"][i], 1), de(res["t"][i] / 60, 3), de(res["core"][i], 2),
                                 de(res["surf"][i], 2), de(res["edge"][i], 2)])
        except OSError as e:
            messagebox.showerror("Export", f"Die Datei konnte nicht geschrieben werden:\n{e}", parent=self)

    def copy_summary(self):
        lines = [APP_TITLE, ""]
        f = self.forge
        if f and f["ok"]:
            b = f["b"]
            lines += ["Schmieden",
                      f"  Werkstoff: {f['mat']['name']}",
                      f"  Rohblock: {geo_label(b['shape'], b['d'], b['L'])}" + (f", r = {nf(b['r'], 1)} mm" if b["r"] > 0 else ""),
                      f"  Schmiedestück: {geo_label(f['shape'], f['d1'], f['L1'])}" + (f", r = {nf(f['r1'], 1)} mm" if f["r1"] > 0 else ""),
                      f"  Materialverlust: {nf(f['loss'], 1)} %",
                      f"  Verschmiedungsgrad: {nf(f['ratio'], 2)} : 1, φ = {nf(f['phi'], 3)}", ""]
        if self.heat and self.heat["res"]["t_reach"] is not None:
            h = self.heat
            r = h["res"]
            lines += ["Aufheizen",
                      f"  {h['geo']['label']}: {geo_label(h['geo']['shape'], h['geo']['d'], h['geo']['L'])}",
                      f"  Stoffwerte: ρ = {nf(h['mat']['rho'], 0)} kg/m³, cp = {nf(h['mat']['cp'], 0)} J/(kg K), λ = {nf(h['mat']['k'], 0)} W/(m K)",
                      f"  T0 = {nf(h['T0'], 0)} °C, Ofen = {nf(h['Tf'], 0)} °C, Ziel Kern = {nf(h['Tz'], 0)} °C",
                      f"  h = {nf(h['hc'], 1)} W/(m² K), ε = {nf(h['eps'], 2)}",
                      f"  Aufheizzeit bis Kern = Ziel: {fmt_time(r['t_reach'])}",
                      f"  Biot-Zahl: {nf(r['Bi'], 3)}, größte Temperaturdifferenz: {nf(r['max_dt'], 1)} K"]
        P = self.plan
        if P and P["res"] and P["res"]["rows"]:
            lines += ["", f"Stichplan ({P['mat']['name']})"]
            b = P["block"]
            lines.append(f"  0: {geo_label(b['shape'], b['d'], b['L'])}")
            for i, r in enumerate(P["res"]["rows"]):
                g = P["geo"][i]
                lines.append(f"  {i + 1}: {geo_label(g['shape'], g['d'], g['L'])}, φ = {nf(g['phi'], 3)}, "
                             f"{r['strokes']} Hübe, F ≈ {nf(r['F'], 2)} MN, Kern {nf(r['T_core'], 0)} °C, "
                             f"Oberfläche {nf(r['T_surf'], 0)} °C, Kontaktzone {nf(r['T_contact'], 0)} °C")
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))

    # ================================================================ QForm-Simulation
    SIM_COLS = (("Stich", 5), ("Schmiedestück", 16), ("Endabstand\nSättel [mm]", 10), ("Hübe", 5),
                ("Bisse ×\nDrehlagen", 9), ("Dreh-\nschritt", 6), ("Vorschub\n[mm]", 9), ("Luft vor\nHub 1 [s]", 9),
                ("Werkstücktemperatur zu Stichbeginn", 34))

    def _pathfield(self, parent, row, label, key, kind, width=40, combo_values=None):
        """Eingabe für Datei- oder Ordnerpfade mit Auswahlknopf; combo_values: zusätzlich Datenbankpfade zur Auswahl."""
        sc, bg = self.sc, C["paper"]
        tk.Label(parent, text=label, bg=bg, font=self.f_base, anchor="w").grid(row=row, column=0, sticky="w",
                                                                               pady=int(3 * sc))
        fr = tk.Frame(parent, bg=bg)
        fr.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=int(3 * sc))
        if combo_values is not None:
            e = ttk.Combobox(fr, textvariable=self.var(key), values=combo_values, width=width)
        else:
            e = ttk.Entry(fr, textvariable=self.var(key), width=width)
        e.pack(side="left", fill="x", expand=True)
        b = ttk.Button(fr, text="…", width=3, command=lambda: self._browse(key, kind))
        b.pack(side="left", padx=(int(4 * sc), 0))
        e._btn = b
        return e

    def _browse(self, key, kind):
        cur = self.v[key].get().strip()
        start = os.path.dirname(cur) if cur and not qform_sim.is_db_path(cur) else self._project_dir()
        if kind == "dir":
            p = filedialog.askdirectory(parent=self, title="Ordner wählen", initialdir=cur or start, mustexist=False)
        else:
            types = qform_sim.QDAT_TYPES if kind == "qdat" else qform_sim.GEOMETRY_TYPES
            p = filedialog.askopenfilename(parent=self, title="Datei wählen", initialdir=start,
                                           filetypes=types)
        if p:
            self.v[key].set(os.path.normpath(p))

    def _build_sim_tab(self):
        sc, bg = self.sc, C["paper"]
        t = self.tab_q
        t.columnconfigure(0, weight=1)
        hd = tk.Frame(t, bg=bg)
        hd.grid(row=0, column=0, sticky="ew", pady=(int(10 * sc), 0))
        tk.Frame(hd, bg=C["ink"], height=max(2, int(2 * sc))).pack(fill="x")
        tk.Label(hd, text="Simulation in QForm aufsetzen", font=self.f_h, bg=bg).pack(anchor="w",
                                                                                      pady=(int(6 * sc), 0))
        self.hint(hd, "Baut den Stichplan über die Python-API von QForm als Projekt auf: je Stich eine Operation mit der "
                      "Hubtabelle (Blows) des Stichplans, Sättel, Werkstück, Temperaturen, Reibung und Presse. Die "
                      "Berechnung läuft als eigener Prozess weiter, auch wenn der Rechner geschlossen wird.", bg,
                  wrap=1000).pack(anchor="w")

        self.qpar = tk.Frame(t, bg=bg)
        self.qpar.grid(row=1, column=0, sticky="ew")
        self.qe = {}

        def fld(sec, row, label, key, unit, width=7):
            e = self.field(sec, row, label, key, unit, bg, width=width)
            self.qe[key] = e
            return e

        # --- QForm und Ablage
        s1 = self._psection(self.qpar, "QForm und Ablage")
        self.qe["qsQDir"] = self._pathfield(s1, 2, "QForm-Verzeichnis", "qsQDir", "dir", width=34)
        self.lb_qdir = self.hint(s1, "", bg, wrap=380)
        self.lb_qdir.grid(row=3, column=0, columnspan=3, sticky="w")
        self.qe["qsOut"] = self._pathfield(s1, 4, "Ausgabeordner", "qsOut", "dir", width=34)
        self.lb_qout = self.hint(s1, "", bg, wrap=380)
        self.lb_qout.grid(row=5, column=0, columnspan=3, sticky="w")
        tk.Label(s1, text="Projektname", bg=bg, font=self.f_base).grid(row=6, column=0, sticky="w", pady=int(3 * sc))
        self.qe["qsName"] = ttk.Entry(s1, textvariable=self.var("qsName"), width=34)
        self.qe["qsName"].grid(row=6, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0))
        self.lb_qname = self.hint(s1, "", bg, wrap=380)
        self.lb_qname.grid(row=7, column=0, columnspan=3, sticky="w")

        # --- Werkstück
        s2 = self._psection(self.qpar, "Werkstück")
        tk.Label(s2, text="Geometrie", bg=bg, font=self.f_base).grid(row=2, column=0, sticky="w")
        self.seg(s2, "qsWpGeo", [("rohblock", "aus Rohblock erzeugen"), ("datei", "aus Datei")], bg).grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=int(2 * sc))
        self.qe["qsWpFile"] = self._pathfield(s2, 3, "Geometriedatei", "qsWpFile", "geo", width=30)
        self.qe["qsWpMat"] = self._pathfield(s2, 4, "Werkstoff", "qsWpMat", "qdat", width=30,
                                             combo_values=qform_db.W_MATERIALS)
        self.hint(s2, "Werkstoff als .qdat-Datei (wird ins Projekt importiert) oder als Pfad der QForm-Datenbank "
                      "(std:w-material/…).", bg, wrap=380).grid(row=5, column=0, columnspan=3, sticky="w")
        tk.Label(s2, text="Temperatur", bg=bg, font=self.f_base).grid(row=6, column=0, sticky="w",
                                                                      pady=(int(4 * sc), 0))
        self.seg(s2, "qsWpTm", [("plan", "wie Stichplan"), ("eigen", "eigener Wert")], bg).grid(
            row=6, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=(int(4 * sc), 0))
        fld(s2, 7, "Werkstücktemperatur", "qsWpT", "°C")
        cb = ttk.Checkbutton(s2, text="Kornentwicklung rechnen", variable=self.var("qsMicro"), onvalue="1",
                             offvalue="0")
        cb.grid(row=8, column=0, columnspan=3, sticky="w", pady=(int(4 * sc), 0))
        fld(s2, 9, "Ausgangskorngröße", "qsGrain", "µm")
        s2.columnconfigure(1, weight=1)

        # --- Sättel und Reibung
        s3 = self._psection(self.qpar, "Sättel und Reibung")
        self.qe["qsToolGeo"] = self._pathfield(s3, 2, "Sattelgeometrie", "qsToolGeo", "geo", width=30)
        self.lb_qtool = self.hint(s3, "", bg, wrap=380)
        self.lb_qtool.grid(row=3, column=0, columnspan=3, sticky="w")
        tk.Label(s3, text="Arbeitsfläche in der Datei", bg=bg, font=self.f_base).grid(row=4, column=0, sticky="w")
        self.seg(s3, "qsToolFace", [("oben", "oben (z max)"), ("unten", "unten (z min)")], bg).grid(
            row=4, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=int(2 * sc))
        self.qe["qsToolMat"] = self._pathfield(s3, 5, "Sattelwerkstoff", "qsToolMat", "qdat", width=30,
                                               combo_values=qform_db.T_MATERIALS)
        tk.Label(s3, text="Temperatur", bg=bg, font=self.f_base).grid(row=6, column=0, sticky="w",
                                                                      pady=(int(4 * sc), 0))
        self.seg(s3, "qsToolTm", [("plan", "wie Stichplan"), ("eigen", "eigener Wert")], bg).grid(
            row=6, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=(int(4 * sc), 0))
        fld(s3, 7, "Satteltemperatur", "qsToolT", "°C")
        tk.Label(s3, text="In jedem Stich", bg=bg, font=self.f_base).grid(row=8, column=0, sticky="w")
        self.seg(s3, "qsToolInh", [("0", "neu setzen"), ("1", "vererben")], bg).grid(
            row=8, column=1, columnspan=2, sticky="ew", padx=(int(8 * sc), 0), pady=int(2 * sc))
        tk.Label(s3, text="Reibung (Schmierstoff)", bg=bg, font=self.f_base).grid(row=9, column=0, sticky="w",
                                                                                  pady=(int(6 * sc), 0))
        self.cb_lub = ttk.Combobox(s3, textvariable=self.var("qsLub"), state="readonly", width=44,
                                   values=[l for l, _p in qform_db.LUBRICANTS] + [qform_db.LUBRICANT_CUSTOM])
        self.cb_lub.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(int(2 * sc), 0))
        self.qe["qsLubPath"] = ttk.Entry(s3, textvariable=self.var("qsLubPath"), width=44)
        self.qe["qsLubPath"].grid(row=11, column=0, columnspan=3, sticky="ew", pady=(int(2 * sc), 0))
        self.lb_qlub = self.hint(s3, "", bg, wrap=380)
        self.lb_qlub.grid(row=12, column=0, columnspan=3, sticky="w")
        s3.columnconfigure(1, weight=1)

        # --- Presse und Manipulator
        s4 = self._psection(self.qpar, "Presse und Manipulator")
        fld(s4, 2, "Hydraulische Presse, Nennkraft", "qsPressF", "MN")
        self.lb_qv = tk.Label(s4, text="", bg=bg, font=self.f_base, anchor="w")
        self.lb_qv.grid(row=3, column=0, columnspan=3, sticky="w", pady=int(3 * sc))
        tk.Label(s4, text="Bissfolge", bg=bg, font=self.f_base).grid(row=4, column=0, sticky="w", pady=(int(4 * sc), 0))
        self.seg(s4, "qsDir", [("manipulator", "vom Manipulator weg"), ("frei", "zum Manipulator hin")], bg).grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(int(2 * sc), 0))
        self.lb_qdir2 = self.hint(s4, "", bg, wrap=360)
        self.lb_qdir2.grid(row=6, column=0, columnspan=3, sticky="w")
        fld(s4, 7, "Zusatzbisse je Stich", "qsExtra", "–")
        fld(s4, 8, "Grifflänge Manipulator", "qsGrip", "mm")
        fld(s4, 9, "Anteil Längung zur Greifseite f", "qsElong", "–")
        fld(s4, 10, "Widerstand axial", "qsResAx", "–")
        fld(s4, 11, "Widerstand quer", "qsResCr", "–")
        ttk.Checkbutton(s4, text="Manipulator stellt nach jedem Hub zurück", variable=self.var("qsGive"),
                        onvalue="1", offvalue="0").grid(row=12, column=0, columnspan=3, sticky="w")
        fld(s4, 13, "Luft über dem Werkstück beim Aufbau", "qsGap", "mm")
        fld(s4, 14, "Rechenzeit je Stich höchstens", "qsMaxH", "h")
        ttk.Checkbutton(s4, text="QForm nach der Berechnung schließen", variable=self.var("qsClose"),
                        onvalue="1", offvalue="0").grid(row=15, column=0, columnspan=3, sticky="w")

        self._sim_sections = [s1, s2, s3, s4]
        self._flow_layout(self.qpar, self._sim_sections)

        # --- Stiche und Wiedererwärmung
        sec, inn = self._section(t, "Stiche und Wiedererwärmung")
        sec.grid(row=2, column=0, sticky="ew", pady=(int(14 * sc), 0))
        self.hint(inn, "Hubtabelle je Stich aus dem Stichplan (Reiter Stichplan). Die Werkstücktemperatur wird von Stich "
                       "zu Stich übernommen; für eine Wiedererwärmung wird der Stich homogen mit der angegebenen "
                       "Temperatur gestartet, vor dem ersten Hub gilt dann die Transportzeit statt der Pause.", bg,
                  wrap=1000).pack(anchor="w")
        self.qtbl = tk.Frame(inn, bg=bg)
        self.qtbl.pack(fill="x", pady=(int(6 * sc), 0))
        self.q_notes = tk.Frame(inn, bg=bg)
        self.q_notes.pack(fill="x")
        self.q_notes.columnconfigure(0, weight=1)

        # --- Simulation starten
        sec2, inn2 = self._section(t, "Simulation")
        sec2.grid(row=3, column=0, sticky="nsew", pady=(int(14 * sc), 0))
        br = tk.Frame(inn2, bg=bg)
        br.pack(fill="x")
        ttk.Button(br, text="Dateien speichern …", command=self.sim_save_files).pack(side="left")
        self.bt_qbuild = ttk.Button(br, text="Stich 1 aufbauen (prüfen)", command=lambda: self.sim_start(True))
        self.bt_qbuild.pack(side="left", padx=(int(8 * sc), 0))
        self.bt_qrun = ttk.Button(br, text="Aufbauen und rechnen", command=lambda: self.sim_start(False))
        self.bt_qrun.pack(side="left", padx=(int(8 * sc), 0))
        self.bt_qstop = ttk.Button(br, text="Abbrechen", command=self.sim_stop, state="disabled")
        self.bt_qstop.pack(side="left", padx=(int(8 * sc), 0))
        ttk.Button(br, text="Ausgabeordner öffnen", command=self.sim_open_folder).pack(side="right")
        self.lb_qstat = tk.Label(inn2, text="Noch keine Simulation gestartet.", font=self.f_sub, bg=bg, fg=C["muted"],
                                 anchor="w", justify="left")
        self.lb_qstat.pack(fill="x", pady=(int(8 * sc), int(2 * sc)))
        self.pb_q = ttk.Progressbar(inn2, mode="determinate", maximum=1.0)
        self.pb_q.pack(fill="x")
        self.q_res = tk.Frame(inn2, bg=bg)
        self.q_res.pack(fill="x", pady=(int(6 * sc), 0))
        lf = tk.Frame(inn2, bg=C["line"])
        lf.pack(fill="both", expand=True, pady=(int(6 * sc), 0))
        mono = tkfont.Font(family="Consolas" if "Consolas" in tkfont.families(self) else "Courier",
                           size=max(8, int(self.f_small.cget("size"))))
        self.q_log = tk.Text(lf, height=14, wrap="none", font=mono, bg=C["field"], fg=C["ink"], relief="flat",
                             padx=int(6 * sc), pady=int(4 * sc), state="disabled")
        ys = ttk.Scrollbar(lf, orient="vertical", command=self.q_log.yview)
        self.q_log.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.q_log.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        self.sim_proc = None
        self.sim_run = None
        self._after_sim = None
        self._after_simpoll = None
        self._q_rows_sig = None

    # ---------- Einstellungen lesen
    def _sim_default_out(self):
        if self.project_path:
            return os.path.join(os.path.dirname(self.project_path), f"QForm_{self._project_name()}")
        docs = os.path.join(os.path.expanduser("~"), "Documents")
        return os.path.join(docs if os.path.isdir(docs) else os.path.expanduser("~"), "Freiformschmieden-Rechner",
                            "QForm")

    def _sim_default_name(self):
        if self.project_path:
            base = self._project_name()
        else:
            base = "Stichplan_" + (self.v["matName"].get().strip() or "Werkstoff")
        return "".join(c if c.isalnum() or c in "-_" else "_" for c in base)

    def _lub_path(self):
        lab = self.v["qsLub"].get()
        for l, p in qform_db.LUBRICANTS:
            if l == lab:
                return p
        return self.v["qsLubPath"].get().strip()

    def read_sim(self):
        """Einstellungen des Reiters; Rückgabe (dict, ok)."""
        g = self.g
        s, ok = {}, True

        def num(key, cond, name):
            nonlocal ok
            val = g(key)
            good = math.isfinite(val) and cond(val)
            if key in self.qe:
                self.mark(self.qe[key], not good)
            ok = ok and good
            s[name] = val
        s["qform_dir"] = self.v["qsQDir"].get().strip() or self._sim_auto_qdir()
        s["out_dir"] = self.v["qsOut"].get().strip() or self._sim_default_out()
        s["name"] = "".join(c if c.isalnum() or c in "-_ " else "_" for c in self.v["qsName"].get().strip()) \
            or self._sim_default_name()
        s["wp_geo"] = self.v["qsWpGeo"].get() or "rohblock"
        s["wp_file"] = self.v["qsWpFile"].get().strip()
        s["wp_mat"] = self.v["qsWpMat"].get().strip()
        s["wp_T_mode"] = self.v["qsWpTm"].get() or "plan"
        s["tool_T_mode"] = self.v["qsToolTm"].get() or "plan"
        if s["wp_T_mode"] == "eigen":
            num("qsWpT", lambda x: 0 < x < 1500, "wp_T")
        else:
            s["wp_T"] = None
            self.qe["qsWpT"].configure(style="Calc.TEntry")
        if s["tool_T_mode"] == "eigen":
            num("qsToolT", lambda x: -50 < x < 1500, "tool_T")
        else:
            s["tool_T"] = None
            self.qe["qsToolT"].configure(style="Calc.TEntry")
        s["micro"] = self.v["qsMicro"].get() == "1"
        if s["micro"]:
            num("qsGrain", lambda x: x > 0, "grain")
        else:
            s["grain"] = g("qsGrain") if math.isfinite(g("qsGrain")) else 100.0
            self.qe["qsGrain"].configure(style="Calc.TEntry")
        s["tool_file"] = self.v["qsToolGeo"].get().strip()
        s["tool_face"] = self.v["qsToolFace"].get() or "oben"
        s["tool_mat"] = self.v["qsToolMat"].get().strip()
        s["tool_inherit"] = self.v["qsToolInh"].get() == "1"
        s["lubricant"] = self._lub_path()
        num("qsPressF", lambda x: x > 0, "press_F")
        s["bissfolge"] = self.v["qsDir"].get() or "manipulator"
        num("qsExtra", lambda x: x >= 0 and abs(x - round(x)) < 1e-9 and x <= 10, "extra_bites")
        s["extra_bites"] = int(round(s["extra_bites"])) if math.isfinite(s["extra_bites"]) else 0
        num("qsGrip", lambda x: x > 0, "grip")
        num("qsElong", lambda x: 0 <= x <= 1, "elong_grip")
        num("qsResAx", lambda x: x >= 0, "res_ax")
        num("qsResCr", lambda x: x >= 0, "res_cross")
        s["give_back"] = self.v["qsGive"].get() == "1"
        num("qsGap", lambda x: x >= 0, "gap")
        num("qsMaxH", lambda x: x > 0, "max_h")
        s["close"] = self.v["qsClose"].get() == "1"
        reheat = []
        for i, r in enumerate(self.plan_rows):
            if i > 0 and r["rw"].get() == "1":
                val = parse(r["rwT"].get())
                good = math.isfinite(val) and 0 < val < 1500
                if "rwT_w" in r and r["rwT_w"].winfo_exists():
                    self.mark(r["rwT_w"], not good)
                ok = ok and good
                reheat.append(val if good else None)
            else:
                reheat.append(None)
        s["reheat"] = reheat
        if self.v["qsLub"].get() == qform_db.LUBRICANT_CUSTOM:
            self.qe["qsLubPath"].grid()
        else:
            self.qe["qsLubPath"].grid_remove()
        self.qe["qsWpFile"].configure(state="normal" if s["wp_geo"] == "datei" else "disabled")
        self.qe["qsWpFile"]._btn.configure(state="normal" if s["wp_geo"] == "datei" else "disabled")
        return s, ok

    # ---------- Darstellung
    def schedule_sim(self):
        if self._after_sim:
            self.after_cancel(self._after_sim)
        self._after_sim = self.after(250, self.run_sim_preview)

    def _sim_plan_ok(self):
        P = self.plan
        return bool(P and P["res"] and P["res"]["rows"] and P["proc"] and len(P["res"]["rows"]) == len(self.plan_rows))

    def run_sim_preview(self):
        """Hinweise, Stichtabelle und Beschriftungen des Reiters auf den Stand bringen."""
        self._after_sim = None
        if not hasattr(self, "qtbl"):
            return
        self.run_plan()
        s, ok = self.read_sim()
        P = self.plan
        found = qform_sim.find_qform_dirs() if sys.platform == "win32" else []
        qd = s["qform_dir"]
        if qform_sim.qform_dir_ok(qd):
            self.lb_qdir.configure(text=f"QForm-API gefunden ({os.path.basename(qd.rstrip(chr(92) + '/'))}).",
                                   fg=C["muted"])
        else:
            alt = f" Gefunden: {', '.join(found)}." if found else ""
            self.lb_qdir.configure(text="QFormAPI.py nicht gefunden (API\\App\\Python)." + alt, fg=C["err"])
        if not self.v["qsQDir"].get().strip():
            self.lb_qdir.configure(text=self.lb_qdir.cget("text") + f" Leer: {qd}")
        self.lb_qout.configure(text="" if self.v["qsOut"].get().strip() else f"Leer: {s['out_dir']}")
        self.lb_qname.configure(text="" if self.v["qsName"].get().strip() else f"Leer: {s['name']}")
        w = qform_sim.die_width_known(s["tool_file"])
        if w:
            bd = P["proc"]["b_die"] if P and P["proc"] else None
            txt = (f"Flachsattel {w[0]:g} mm breit in Stabachse, ebene Arbeitsfläche {w[1]:g} mm (Kantenradius "
                   f"{w[2]:g} mm)." + (f" Sattelbreite im Stichplan: {bd:g} mm." if bd else ""))
            self.lb_qtool.configure(text=txt, fg=C["warn"] if bd and bd > w[1] + 0.5 else C["muted"])
        else:
            self.lb_qtool.configure(text="Sattelbreite in Stabachse muss der Sattelbreite im Stichplan entsprechen.",
                                    fg=C["muted"])
        self.lb_qlub.configure(text=f"QForm-Datenbank: {s['lubricant'] or '–'}. Der Schmierstoff legt in QForm "
                                    "Reibung und Wärmeübergang zum Werkzeug fest.")
        if P and P["proc"]:
            self.lb_qv.configure(text=f"Hubgeschwindigkeit aus dem Stichplan: {nf(P['proc']['v'], 0)} mm/s")
            Tw = P["proc"]["T_start"]
            Tt = P["proc"]["T_die"]
            if s["wp_T_mode"] == "plan":
                self._set_calc(self.qe["qsWpT"], self.v["qsWpT"], fe(Tw, 0), True)
            if s["tool_T_mode"] == "plan":
                self._set_calc(self.qe["qsToolT"], self.v["qsToolT"], fe(Tt, 0), True)
        if s["bissfolge"] == "manipulator":
            self.lb_qdir2.configure(text="Erster Biss am gegriffenen Ende, Vorschub s·(1 + (λ−1)·(1 − f)), das "
                                         "Werkstück fährt zum Manipulator. f: Anteil der Längung eines Bisses, der die "
                                         "Greifseite verschiebt (0 = starrer Manipulator, Vorschub s·λ wie im "
                                         "Stichplan; in QForm mit Rückstellung etwa 0,5).")
        else:
            self.lb_qdir2.configure(text="Erster Biss am freien Ende, Vorschub s·(1 + (λ−1)·f), das Werkstück fährt "
                                         "vom Manipulator weg. f: Anteil der Längung eines Bisses, der die Greifseite "
                                         "verschiebt (0 = starrer Manipulator, Vorschub s wie Stichplan V1/V3; in "
                                         "QForm mit Rückstellung etwa 0,5).")
        self._render_sim_table(s)
        self._render_sim_notes(s, ok)
        self.sf_q.refit()

    def _render_sim_table(self, s):
        sc, bg = self.sc, C["paper"]
        P = self.plan
        rows = P["res"]["rows"] if P and P["res"] else []
        sig = (len(self.plan_rows), len(rows))
        if sig != self._q_rows_sig:
            for w in self.qtbl.winfo_children():
                w.destroy()
            self._q_rows_sig = sig
            for j, (h, width) in enumerate(self.SIM_COLS):
                tk.Label(self.qtbl, text=h, font=self.f_small, fg=C["muted"], bg=bg, justify="center", width=width,
                         anchor="s").grid(row=0, column=j, sticky="sew", padx=int(1 * sc))
            tk.Frame(self.qtbl, bg=C["line_strong"], height=1).grid(row=1, column=0, columnspan=len(self.SIM_COLS),
                                                                    sticky="ew")
            self._q_cells = []
            for i, r in enumerate(self.plan_rows):
                cells = []
                for j, (_h, width) in enumerate(self.SIM_COLS[:-1]):
                    lab = tk.Label(self.qtbl, text="", font=self.f_val if j == 2 else self.f_base, bg=bg,
                                   width=width, anchor="w" if j == 1 else "e")
                    lab.grid(row=2 + i, column=j, sticky="ew", padx=int(1 * sc), pady=int(1 * sc))
                    cells.append(lab)
                fr = tk.Frame(self.qtbl, bg=bg)
                fr.grid(row=2 + i, column=len(self.SIM_COLS) - 1, sticky="w", padx=(int(8 * sc), 0))
                if i == 0:
                    lab = tk.Label(fr, text="", font=self.f_base, bg=bg, fg=C["muted"])
                    lab.pack(side="left")
                    cells.append(lab)
                else:
                    disp = tk.StringVar(self)
                    r["rw_disp"] = disp
                    cbx = ttk.Combobox(fr, textvariable=disp, state="readonly", width=30,
                                       values=["aus vorherigem Stich übernehmen", "Wiedererwärmen auf"])
                    cbx.pack(side="left")

                    def on_sel(_e=None, row=r, dv=disp):
                        row["rw"].set("1" if dv.get().startswith("Wieder") else "0")
                        if row["rw"].get() == "1" and not row["rwT"].get().strip():
                            row["rwT"].set(fe(self.plan["proc"]["T_start"], 0) if self.plan and self.plan["proc"]
                                           else "")
                    cbx.bind("<<ComboboxSelected>>", on_sel)
                    e = ttk.Entry(fr, textvariable=r["rwT"], width=6, justify="right")
                    e.pack(side="left", padx=(int(6 * sc), int(3 * sc)))
                    tk.Label(fr, text="°C", font=self.f_small, fg=C["muted"], bg=bg).pack(side="left")
                    r["rwT_w"] = e
                    cells.append(e)
                self._q_cells.append(cells)
        ok_plan = self._sim_plan_ok()
        try:
            cfg, _notes, prev = qform_sim.build_config(self.plan, s, VERSION) if ok_plan and all(
                v is not None for k, v in s.items() if k in ("press_F",)) else (None, None, [])
        except Exception:
            prev = []
        for i, cells in enumerate(getattr(self, "_q_cells", [])):
            d = prev[i] if i < len(prev) else None
            if d:
                vals = (str(d["nr"]), f"{SHAPES[d['shape']]['pre']} {fdim(d['d'])} × {nf(d['L'], 0)}", fdim(d["d"]),
                        str(d["blows"]), f"{d['n_b']} × {d['n_rot']}", f"{nf(d['rot'], 0)}°",
                        ("+" if d["feed"] > 0 else "") + nf(d["feed"], 1) if d["n_b"] > 1 else "–", nf(d["air"], 0))
            else:
                vals = (str(i + 1),) + ("–",) * 7
            for j, v in enumerate(vals):
                cells[j].configure(text=v)
            if i == 0:
                Tw = None
                if self.plan and self.plan["proc"]:
                    Tw = self.plan["proc"]["T_start"] if s["wp_T_mode"] == "plan" else s.get("wp_T")
                cells[-1].configure(text=f"Start mit {nf(Tw, 0)} °C (Werkstücktemperatur)" if Tw else "–")
            else:
                r = self.plan_rows[i]
                want = "Wiedererwärmen auf" if r["rw"].get() == "1" else "aus vorherigem Stich übernehmen"
                if "rw_disp" in r and r["rw_disp"].get() != want:
                    r["rw_disp"].set(want)
                cells[-1].configure(state="normal" if r["rw"].get() == "1" else "disabled")

    def _render_sim_notes(self, s, ok):
        p = self.q_notes
        for w in p.winfo_children():
            w.destroy()
        notes = []
        if not self._sim_plan_ok():
            notes.append(("err", "Der Stichplan ist noch nicht vollständig gerechnet (Reiter Stichplan prüfen)."))
        else:
            try:
                _cfg, n2, _prev = qform_sim.build_config(self.plan, s, VERSION)
                notes += n2
            except Exception as ex:
                notes.append(("err", f"Eingaben unvollständig ({ex})."))
        if not ok:
            notes.insert(0, ("err", "Eingaben prüfen (rot markierte Felder)."))
        if sys.platform != "win32":
            notes.append(("info", "QForm läuft nur unter Windows; hier können nur die Dateien gespeichert werden."))
        k = 0
        for kind in ("err", "warn", "info"):
            texts = [t for kk, t in notes if kk == kind]
            if texts:
                self._note(p, kind, "\n".join(texts), k)
                lab = p.grid_slaves(row=k)[0].winfo_children()[0]
                lab.configure(wraplength=max(400, self.sf_q.winfo_width() - int(60 * self.sc)), pady=int(4 * self.sc))
                k += 1
        self._sim_errors = [t for kk, t in notes if kk == "err"]

    # ---------- Aktionen
    def _sim_config(self):
        """Konfiguration erzeugen; bei Fehlern Meldung und None."""
        self.run_sim_preview()
        if self._sim_errors:
            messagebox.showerror("QForm-Simulation", "Die Simulation kann noch nicht aufgesetzt werden:\n\n• "
                                 + "\n• ".join(self._sim_errors), parent=self)
            return None
        s, ok = self.read_sim()
        if not ok:
            return None
        cfg, _notes, _prev = qform_sim.build_config(self.plan, s, VERSION)
        return cfg

    def sim_save_files(self):
        cfg = self._sim_config() if self._sim_plan_ok() else None
        if cfg is None:
            return
        folder = filedialog.askdirectory(parent=self, title="Ordner für die Simulationsdateien",
                                         initialdir=cfg["ausgabe_dir"] if os.path.isdir(cfg["ausgabe_dir"])
                                         else self._project_dir(), mustexist=False)
        if not folder:
            return
        cfg["ausgabe_dir"] = os.path.normpath(folder)
        try:
            files = qform_sim.write_package(folder, cfg)
        except OSError as e:
            messagebox.showerror("QForm-Simulation", f"Die Dateien konnten nicht geschrieben werden:\n{e}", parent=self)
            return
        messagebox.showinfo("QForm-Simulation", f"Gespeichert in {folder}:\n" + "\n".join(
            "  " + os.path.basename(f) for f in files) + "\n\nStart mit Simulation_berechnen.bat oder "
            "Simulation_Aufbau_pruefen.bat (Python 3 und QForm mit API erforderlich).", parent=self)

    def sim_start(self, nur_aufbau):
        if self.sim_proc is not None and self.sim_proc.poll() is None:
            messagebox.showinfo("QForm-Simulation", "Es läuft bereits eine Simulation.", parent=self)
            return
        if sys.platform != "win32":
            messagebox.showinfo("QForm-Simulation", "QForm läuft nur unter Windows. Mit „Dateien speichern …“ lässt "
                                                    "sich die Simulation auf einem QForm-Rechner starten.", parent=self)
            return
        cfg = self._sim_config() if self._sim_plan_ok() else None
        if cfg is None:
            if not self._sim_plan_ok():
                messagebox.showerror("QForm-Simulation", "Der Stichplan ist noch nicht vollständig gerechnet.",
                                     parent=self)
            return
        n = sum(len(x["blows"]) for x in cfg["stiche"])
        if not nur_aufbau and not messagebox.askyesno(
                "QForm-Simulation", f"{len(cfg['stiche'])} Stiche mit {n} Hüben aufbauen und rechnen?\n\nQForm belegt "
                                    "dabei einen Lizenzplatz, die Berechnung kann mehrere Stunden dauern. Andere "
                                    "QForm-Fenster vorher schließen.", parent=self):
            return
        out = cfg["ausgabe_dir"]
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        prot = os.path.join(out, "Protokolle")
        try:
            os.makedirs(prot, exist_ok=True)
            qform_sim.write_package(out, cfg)
        except OSError as e:
            messagebox.showerror("QForm-Simulation", f"Ausgabeordner nicht beschreibbar:\n{e}", parent=self)
            return
        base = os.path.join(prot, f"{cfg['projektname']}_{stamp}")
        run = dict(cfg=os.path.join(out, "simulation.json"), log=base + ".log", status=base + "_status.json",
                   stop=os.path.join(out, "STOP"), out=out, nur_aufbau=nur_aufbau, start=stamp)
        try:
            if os.path.exists(run["stop"]):
                os.remove(run["stop"])
            self.sim_proc = qform_sim.start_run(run["cfg"], nur_aufbau, run["log"], run["status"], run["stop"])
        except OSError as e:
            messagebox.showerror("QForm-Simulation", f"Der Simulationsprozess konnte nicht gestartet werden:\n{e}",
                                 parent=self)
            return
        self.sim_run = run
        self.v["qsLast"].set(json.dumps({k: run[k] for k in ("log", "status", "stop", "out", "start")}))
        self._log_size = -1
        self.lb_qstat.configure(text="QForm wird gestartet …", fg=C["ink"])
        self.pb_q.configure(value=0.0)
        for w in self.q_res.winfo_children():
            w.destroy()
        self._sim_buttons(True)
        self._sim_poll()

    def sim_stop(self):
        run = self.sim_run
        if not run:
            return
        if not messagebox.askyesno("QForm-Simulation", "Laufende Simulation abbrechen? Das Projekt wird mit dem "
                                                       "erreichten Stand gespeichert.", parent=self):
            return
        try:
            with open(run["stop"], "w", encoding="utf-8") as fh:
                fh.write("stop\n")
            self.lb_qstat.configure(text="Abbruch angefordert …", fg=C["warn"])
        except OSError as e:
            messagebox.showerror("QForm-Simulation", f"Abbruch nicht möglich:\n{e}", parent=self)

    def sim_open_folder(self):
        s, _ok = self.read_sim()
        p = s["out_dir"]
        if not os.path.isdir(p):
            messagebox.showinfo("QForm-Simulation", f"Der Ausgabeordner existiert noch nicht:\n{p}", parent=self)
            return
        try:
            if sys.platform == "win32":
                os.startfile(p)  # noqa: S606
            else:
                subprocess.Popen(["xdg-open", p])
        except OSError:
            pass

    def _sim_buttons(self, running):
        self.bt_qstop.configure(state="normal" if running else "disabled")
        self.bt_qbuild.configure(state="disabled" if running else "normal")
        self.bt_qrun.configure(state="disabled" if running else "normal")

    def _sim_poll(self):
        """Protokoll und Status des laufenden Prozesses anzeigen."""
        self._after_simpoll = None
        run = self.sim_run
        if not run:
            return
        st = None
        try:
            with open(run["status"], "r", encoding="utf-8") as fh:
                st = json.load(fh)
        except (OSError, ValueError):
            pass
        final = ("fertig", "abgebrochen", "fehler")
        if self.sim_proc is not None:
            running = self.sim_proc.poll() is None
        else:                                   # Lauf aus einer früheren Sitzung: Status muss frisch sein
            try:
                fresh = time.time() - os.path.getmtime(run["status"]) < 300
            except OSError:
                fresh = False
            running = bool(st) and st.get("zustand") not in final and fresh
        try:
            size = os.path.getsize(run["log"])
            if size != getattr(self, "_log_size", -1):
                self._log_size = size
                with open(run["log"], "r", encoding="utf-8", errors="replace") as fh:
                    text = fh.read()
                lines = text.splitlines()
                at_end = self.q_log.yview()[1] > 0.98
                self.q_log.configure(state="normal")
                self.q_log.delete("1.0", "end")
                self.q_log.insert("end", "\n".join(lines[-600:]))
                self.q_log.configure(state="disabled")
                if at_end or not running:
                    self.q_log.see("end")
        except OSError:
            pass
        if st:
            z = st.get("zustand", "")
            stiche = max(1, int(st.get("stiche") or 1))
            k = int(st.get("stich") or 0)
            hub, huebe = int(st.get("hub") or 0), max(1, int(st.get("huebe") or 1))
            frac = ((k - 1) + min(1.0, hub / huebe)) / stiche if k else 0.0
            if z == "fertig":
                frac = 1.0
            self.pb_q.configure(value=max(0.0, min(1.0, frac)))
            col = {"fehler": C["err"], "abgebrochen": C["warn"], "fertig": C["cold"]}.get(z, C["ink"])
            txt = {"startet": "QForm wird gestartet …", "aufbau": st.get("meldung", "Aufbau …"),
                   "rechnet": st.get("meldung", "Berechnung läuft …"),
                   "fertig": ("Aufbau fertig, QForm bleibt zur Kontrolle offen." if st.get("nur_aufbau")
                              else "Berechnung fertig."),
                   "abgebrochen": "Abgebrochen. Das Projekt ist mit dem erreichten Stand gespeichert.",
                   "fehler": "Fehler: " + str(st.get("meldung", ""))}.get(z, st.get("meldung", ""))
            if z in ("fertig", "abgebrochen", "fehler"):
                txt += f"\nProjekt: {st.get('projekt', '')}"
            self.lb_qstat.configure(text=txt, fg=col)
            self._render_sim_results(st.get("ergebnisse") or [])
        if running:
            self._sim_buttons(True)
            self._after_simpoll = self.after(1500, self._sim_poll)
        else:
            if not st or st.get("zustand") not in final:
                self.lb_qstat.configure(text="Kein laufender Simulationsprozess (keine Abschlussmeldung im Status). "
                                             "Protokoll prüfen.", fg=C["err"])
            self._sim_buttons(False)
            self.sim_proc = None

    def _render_sim_results(self, erg):
        p = self.q_res
        sig = json.dumps([(e.get("stich"), e.get("laenge_ist"), e.get("f_max")) for e in erg])
        if getattr(p, "_sig", None) == sig:
            return
        p._sig = sig
        for w in p.winfo_children():
            w.destroy()
        if not erg:
            return
        bg, sc = C["paper"], self.sc
        heads = ("Stich", "Länge Stichplan\n[mm]", "Länge QForm\n[mm]", "Querschnitt Mitte\nQForm [mm]",
                 "Presskraft Stichplan\n[MN]", "Presskraft QForm\n[MN]", "Rechenzeit\n[min]")
        for j, h in enumerate(heads):
            tk.Label(p, text=h, bg=C["panel"], font=self.f_small, padx=int(6 * sc), justify="center").grid(
                row=0, column=j, sticky="nsew")
        for i, e in enumerate(erg):
            qm = e.get("querschnitt_mitte") or ["–", "–"]
            vals = (str(e.get("stich")), nf(e.get("laenge_stichplan"), 0), nf(e.get("laenge_ist"), 0),
                    f"{nf(qm[0], 1) if isinstance(qm[0], (int, float)) else qm[0]} × "
                    f"{nf(qm[1], 1) if isinstance(qm[1], (int, float)) else qm[1]}",
                    nf(e.get("f_stichplan_MN"), 2), nf(e.get("f_max"), 2) if e.get("f_max") is not None else "–",
                    nf(e.get("rechenzeit_min"), 1))
            for j, v in enumerate(vals):
                tk.Label(p, text=v, bg=bg, font=self.f_base, anchor="e", padx=int(6 * sc)).grid(
                    row=i + 1, column=j, sticky="ew")

    def _sim_auto_qdir(self):
        if not hasattr(self, "_qdirs_found"):
            self._qdirs_found = qform_sim.find_qform_dirs() if sys.platform == "win32" else []
        return self._qdirs_found[0] if self._qdirs_found else qform_sim.DEFAULT_QFORM_DIR

    def _sim_init(self):
        """Nach dem Start: Protokoll des letzten Laufs anzeigen; läuft er noch, wird weiter mitgelesen."""
        try:
            last = json.loads(self.v["qsLast"].get() or "{}")
        except ValueError:
            last = {}
        if last.get("log") and os.path.isfile(last["log"]) and last.get("status"):
            self.sim_run = dict(last)
            self._log_size = -1
            self.sim_proc = None
            self._sim_poll()

    def show_model(self):
        top = tk.Toplevel(self)
        top.title("Rechenmodell und Annahmen")
        top.configure(bg=C["paper"])
        top.transient(self)
        sbar = ttk.Scrollbar(top, orient="vertical")
        txt = tk.Text(top, wrap="word", width=90, height=34, bg=C["paper"], fg=C["ink"], font=self.f_base,
                      yscrollcommand=sbar.set,
                      relief="flat", padx=int(16 * self.sc), pady=int(12 * self.sc))
        txt.insert("1.0", MODEL_TEXT)
        for head in ("Endlänge", "Aufheizen", "Randbedingung", "Nicht enthalten", "Stoffwerte",
                     "Stichplan: Geometrie", "Stichplan: Temperatur", "Stichplan: QForm-Export",
                     "Stichplan: Fließspannung und Kraft",
                     "Stichplan: Nicht enthalten", "QForm-Simulation"):
            idx = txt.search(head + "\n", "1.0", stopindex="end")
            if idx:
                txt.tag_add("h", idx, f"{idx} lineend")
        txt.tag_configure("h", font=self.f_h, spacing1=int(8 * self.sc))
        txt.configure(state="disabled")
        sbar.configure(command=txt.yview)
        ttk.Button(top, text="Schließen", command=top.destroy).pack(side="bottom", pady=int(10 * self.sc))
        sbar.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)


def main():
    if "--qform-run" in sys.argv:                    # Simulationsprozess ohne Oberfläche (aus dem Reiter gestartet)
        i = sys.argv.index("--qform-run")
        sys.exit(qform_runner.main(sys.argv[i + 1:]))
    if sys.platform == "win32":
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    app = App()
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args and os.path.isfile(args[0]):
        app.after(300, lambda: app.open_project(args[0]))
    app.mainloop()


if __name__ == "__main__":
    main()
