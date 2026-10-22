# -*- coding: utf-8 -*-
"""
Export des Stichplans als QForm-Hubtabelle (Registerkarte „Blows“, Import aus Excel)
===================================================================================
Format wie die QForm-Vorlage: ein Blatt „Table“, Zeile 1 Bezeichnungen, Zeile 2 Einheiten, ab Zeile 3 ein Hub je Zeile.

    Blow | CoolingInAir | CoolingInTool | Displacement | RotationAngle
    #    | s            | s             | mm           | degree

Belegung (QForm wendet alle Größen vor dem jeweiligen Hub an):
  CoolingInAir   erster Hub des Stichplans: Transport zur Presse; erster Hub jedes weiteren Stichs: Pause zwischen
                 den Stichen; alle übrigen Hübe: 0
  CoolingInTool  Nebenzeit je Hub (Werkstück liegt auf dem Untersattel), vor jedem Hub
  Displacement   Vorschub vor dem ersten Hub jedes neuen Bisses = Bisslänge × Streckgrad des Stichs (A_vorher/A_nachher);
                 so trifft der Sattel nach der Längung genau auf ungeschmiedeten Werkstoff (minimale Bisszahl).
                 Optional fährt der erste Hub jedes weiteren Stichs an den Stichanfang zurück.
  RotationAngle  180°/Drehlagen vor jedem Hub außer dem ersten Hub eines Stichs (je Biss alle Drehlagen)
"""
import csv

HEAD = ("Blow", "CoolingInAir", "CoolingInTool", "Displacement", "RotationAngle")
UNITS = ("#", "s", "s", "mm", "degree")


def _num(x, nd):
    v = round(float(x), nd)
    if v == 0:
        return 0
    return int(v) if v == int(v) else v


def build_blows(rows, proc, sign=-1, return_to_start=True):
    """rows: Ergebnisse je Stich aus dem Stichplan (n_b, n_rot, feed, rot_step).
    Rückgabe: Liste je Stich, darin je Hub (CoolingInAir, CoolingInTool, Displacement, RotationAngle)."""
    out = []
    travel_prev = 0.0
    for k, r in enumerate(rows):
        blows = []
        for j in range(int(r["n_b"])):
            for q in range(int(r["n_rot"])):
                if j == 0 and q == 0:
                    air = proc["t_transfer"] if k == 0 else proc["t_pause"]
                    disp = -sign * travel_prev if (k > 0 and return_to_start) else 0.0
                    rot = 0.0
                else:
                    air = 0.0
                    disp = sign * r["feed"] if q == 0 else 0.0
                    rot = r["rot_step"]
                blows.append((_num(air, 2), _num(proc["t_aux"], 2), _num(disp, 1), _num(rot, 2)))
        travel_prev = (int(r["n_b"]) - 1) * r["feed"]
        out.append(blows)
    return out


def table_rows(blows):
    """Hubliste -> Tabellenzeilen mit fortlaufender Hubnummer."""
    return [(i + 1,) + tuple(b) for i, b in enumerate(blows)]


def write_xlsx(path, blows):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Table"
    ws.append(list(HEAD))
    ws.append(list(UNITS))
    for row in table_rows(blows):
        ws.append(list(row))
    wb.save(path)


def write_csv(path, blows, decimal="."):
    def fmt(v):
        s = repr(v) if isinstance(v, float) else str(v)
        return s.replace(".", ",") if decimal == "," else s
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        wr = csv.writer(fh, delimiter=";")
        wr.writerow(HEAD)
        wr.writerow(UNITS)
        for row in table_rows(blows):
            wr.writerow([fmt(v) for v in row])
