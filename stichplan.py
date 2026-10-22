# -*- coding: utf-8 -*-
"""
Stichplan für das Freiformschmieden (Recken mit Flachsätteln)
=============================================================
* Geometrie je Stich: Maß oder Länge vorgeben, das jeweils andere aus der Volumenkonstanz
* Fließspannung: Sellars-Tegart (Zener-Hollomon) oder Hensel-Spittel (T in °C)
* Temperatur je Stich für eine Bezugsscheibe in Stabmitte:
    - Wärmeleitung im Querschnitt: 1D-Zylinder mit gleichem Verhältnis Volumen/Oberfläche (R = 2A/U)
    - Werkzeugkontakt: zwei halbunendliche Körper mit Kontaktleitwert h_K (geschlossene Lösung),
      Kontaktzeit aus Hubweg und Hubgeschwindigkeit
    - Umformwärme: beta * k_f * dphi / (eta_F * rho * c), gleichmäßig im Querschnitt
    - Luft: Konvektion und Strahlung
* Presskraft: k_f * Kontaktfläche * Formfaktor (Hill / Siebel)
* Hubfolge wie in QForm (Blows): je Biss alle Drehlagen, dann Vorschub zum nächsten Biss. Vor jedem Hub liegt
  das Werkstück während der Nebenzeit auf dem Untersattel (QForm: CoolingInTool), Kontaktleitwert h_rest.
  Vorschub = Bisslänge × Streckgrad des Stichs, damit der Sattel nach der Längung auf ungeschmiedetem Werkstoff landet.

Nur Standardbibliothek.
"""
import math

SQ2 = math.sqrt(2.0)
SIGMA = 5.670374419e-8
R_GAS = 8.314462618
BETA_TQ = 0.9          # Taylor-Quinney-Koeffizient
T_AMB = 20.0           # Umgebungstemperatur °C
RC_DIE = 3.6e6         # rho*c Warmarbeitsstahl, J/(m³ K)

DEFAULT_NROT = {"vier": 2, "acht": 4, "rund": 4}


# ---------------------------------------------------------------- Geometrie (scharfkantig)
def area(shape, d):
    if shape == "rund":
        return math.pi * d * d / 4.0
    if shape == "vier":
        return d * d
    return 2.0 * (SQ2 - 1.0) * d * d


def dim_from_area(shape, A):
    if shape == "rund":
        return math.sqrt(4.0 * A / math.pi)
    if shape == "vier":
        return math.sqrt(A)
    return math.sqrt(A / (2.0 * (SQ2 - 1.0)))


def perimeter(shape, d):
    if shape == "rund":
        return math.pi * d
    if shape == "vier":
        return 4.0 * d
    return 8.0 * d * (SQ2 - 1.0)


def height_before(prev_shape, prev_d, new_shape):
    """Ausdehnung des vorherigen Querschnitts in den Druckrichtungen des neuen Stichs (gemittelt)."""
    if prev_shape == "rund":
        return prev_d
    if new_shape == "vier":
        return prev_d                              # Vierkant: a, Achtkant: SW (Flächen bei 0°/90°)
    if new_shape == "acht":
        return prev_d * (1.0 + SQ2) / 2.0 if prev_shape == "vier" else prev_d   # Flächen und Ecken
    return perimeter(prev_shape, prev_d) / math.pi  # Rund: mittlerer Durchmesser (Umfang/π)


def bite_layout(L_prev, h_prev, ratio, proc):
    """Bisse eines Stichs. Rückgabe: Anzahl Bisse, Bisslänge (regulär), letzter Biss, Vorschub, größte Bisslänge.
    Modus 'min': größte Bisslänge = Sattelbreite (minimale Bisszahl); 'const': konstante Bisslänge für alle Stiche;
    'sbh': aus s_B/h. Die Bisslänge ist immer höchstens die Sattelbreite."""
    mode = proc.get("bite_mode", "min")
    if mode == "const":
        s_max = min(proc["sB_const"], proc["b_die"])
    elif mode == "sbh":
        s_max = min(proc["sB_rel"] * h_prev, proc["b_die"])
    else:
        s_max = proc["b_die"]
    n_b = max(1, int(math.ceil(L_prev / s_max - 1e-9)))
    s = min(s_max, L_prev)
    s_last = L_prev - (n_b - 1) * s
    # kleiner Rest (< 1/4 Bisslänge) wird dem letzten Biss zugeschlagen, wenn dieser dann noch unter den Sattel passt
    if n_b > 1 and s_last < 0.25 * s and s + s_last <= proc["b_die"] + 1e-9:
        n_b -= 1
        s_last += s
    return n_b, s, s_last, s * ratio, s_max


def contact_width_auto(new_shape, new_d, h_prev):
    """Kontaktbreite quer zur Stabachse je Sattel am Hubende (Flachsättel)."""
    if new_shape == "vier":
        return new_d
    if new_shape == "acht":
        return new_d * (SQ2 - 1.0)
    delta = max(h_prev - new_d, 0.0) / 2.0          # Abplattung je Seite
    b = 2.0 * math.sqrt(max(delta * (h_prev - delta), 0.0))
    return min(max(b, 0.1 * new_d), 0.8 * new_d)


# ---------------------------------------------------------------- Fließspannung
def kf_sellars(T, phid, p):
    """Sellars-Tegart, stationär. T in °C, phid in 1/s, p: alpha [1/MPa], n, Q [kJ/mol], lnA. Ergebnis MPa."""
    TK = max(T, 20.0) + 273.15
    lnZ = math.log(max(phid, 1e-6)) + p["Q"] * 1e3 / (R_GAS * TK)
    x = (lnZ - p["lnA"]) / p["n"]
    if x > 700:
        return float("inf")
    return math.asinh(math.exp(x)) / p["alpha"]


def kf_hensel(T, phi, phid, p):
    """Hensel-Spittel (9 Parameter, T in °C). Ergebnis MPa."""
    phi = max(phi, 1e-3)
    phid = max(phid, 1e-6)
    T = max(T, 1.0)
    ln = (math.log(p["A"]) + p["m1"] * T + p["m9"] * math.log(T) + p["m2"] * math.log(phi) + p["m4"] / phi
          + p["m5"] * T * math.log(1.0 + phi) + p["m7"] * phi + p["m3"] * math.log(phid)
          + p["m8"] * T * math.log(phid))
    return math.exp(min(ln, 700.0))


def kf_mean(model, T, phi0, phi1, phid):
    """Mittlere Fließspannung über das Umformgradintervall [phi0, phi1]."""
    if model["type"] == "st":
        return kf_sellars(T, phid, model)
    a, b = max(phi0, 0.02), max(phi1, 0.03)
    if b <= a:
        return kf_hensel(T, b, phid, model)
    n = 8
    hstep = (b - a) / n
    s = kf_hensel(T, a, phid, model) + kf_hensel(T, b, phid, model)
    for i in range(1, n):
        s += (4 if i % 2 else 2) * kf_hensel(T, a + i * hstep, phid, model)
    return s * hstep / 3.0 / (b - a)


def st_lnA_from_anchor(alpha, n, Q, T_ref, sigma_ref, phid_ref=1.0):
    """lnA so, dass kf(T_ref, phid_ref) = sigma_ref."""
    TK = T_ref + 273.15
    lnZ = math.log(phid_ref) + Q * 1e3 / (R_GAS * TK)
    return lnZ - n * math.log(math.sinh(alpha * sigma_ref))


def _solve(M, v):
    n = len(v)
    A = [row[:] + [v[i]] for i, row in enumerate(M)]
    for c in range(n):
        piv = max(range(c, n), key=lambda r: abs(A[r][c]))
        A[c], A[piv] = A[piv], A[c]
        for r in range(c + 1, n):
            f = A[r][c] / A[c][c]
            for k in range(c, n + 1):
                A[r][k] -= f * A[c][k]
    x = [0.0] * n
    for r in range(n - 1, -1, -1):
        x[r] = (A[r][n] - sum(A[r][k] * x[k] for k in range(r + 1, n))) / A[r][r]
    return x


def hs_from_st(st, T_lo, T_hi):
    """Hensel-Spittel-Startwerte (A, m1, m3, m8) per Ausgleichsrechnung an ein Sellars-Tegart-Modell."""
    rows, ys = [], []
    for i in range(9):
        T = T_lo + (T_hi - T_lo) * i / 8.0
        for phid in (0.01, 0.1, 1.0, 10.0):
            lp = math.log(phid)
            rows.append([1.0, T, lp, T * lp])
            ys.append(math.log(kf_sellars(T, phid, st)))
    M = [[sum(r[i] * r[j] for r in rows) for j in range(4)] for i in range(4)]
    v = [sum(r[i] * y for r, y in zip(rows, ys)) for i in range(4)]
    c = _solve(M, v)
    return dict(A=math.exp(c[0]), m1=c[1], m2=0.0, m3=c[2], m4=0.0, m5=0.0, m7=0.0, m8=c[3], m9=0.0)


# ---------------------------------------------------------------- Werkzeugkontakt
def erfcx(x):
    """exp(x²)·erfc(x), auch für große x stabil."""
    if x < 6.0:
        return math.exp(x * x) * math.erfc(x)
    x2 = x * x
    return (1.0 - 1.0 / (2 * x2) + 3.0 / (4 * x2 * x2) - 15.0 / (8 * x2 ** 3)) / (x * math.sqrt(math.pi))


def contact_two_halfspaces(T1, T2, h, e1, e2, t):
    """Zwei halbunendliche Körper (Werkstück T1, Werkzeug T2) mit Kontaktleitwert h für die Dauer t.
    e = sqrt(lambda·rho·c). Rückgabe: Oberflächentemperatur Werkstück am Ende [°C],
    entzogene Wärme je Kontaktfläche [J/m²]."""
    dT = T1 - T2
    if t <= 0 or h <= 0:
        return T1, 0.0
    g = h * (1.0 / e1 + 1.0 / e2)
    s = g * math.sqrt(t)
    ex = erfcx(s)
    Ts = T1 - dT * (e2 / (e1 + e2)) * (1.0 - ex)
    E = h * dT / (g * g) * (ex - 1.0 + 2.0 * s / math.sqrt(math.pi))
    return Ts, E


# ---------------------------------------------------------------- Wärmeleitung im Querschnitt
class Radial:
    """1D-Zylinder (Knoten bei r = i/N·R), implizit, mit Volumenquelle und gemischter Randbedingung."""

    def __init__(self, N, T0):
        self.N = N
        self.T = [T0] * (N + 1)
        xi = [i / N for i in range(N + 1)]
        self.v = []
        for i in range(N + 1):
            a = max(0.0, xi[i] - 0.5 / N)
            b = min(1.0, xi[i] + 0.5 / N)
            self.v.append((b * b - a * a) / 2.0)       # Volumen je (R² · rad · m)
        self.f = [(i + 0.5) / N for i in range(N)]      # Flächen je (R · rad · m)
        self.vt = sum(self.v)

    def mean(self):
        return sum(v * t for v, t in zip(self.v, self.T)) / self.vt

    def step(self, R, k, rc, dt, h_air, f_air, q_die, S, eps):
        """Ein impliziter Zeitschritt. q_die: Wärmestrom zum Werkzeug, auf die gesamte Oberfläche gemittelt [W/m²].
        S: Volumenquelle [W/m³]. Luft: Konvektion h_air und Strahlung eps auf dem Anteil f_air der Oberfläche."""
        N, T = self.N, self.T
        dr = R / N
        Ts = T[N]
        K, Ka = Ts + 273.15, T_AMB + 273.15
        hr = eps * SIGMA * (K * K + Ka * Ka) * (K + Ka)
        heff = f_air * (h_air + hr)
        a = [0.0] * (N + 1)
        b = [0.0] * (N + 1)
        c = [0.0] * (N + 1)
        d = [0.0] * (N + 1)
        for i in range(N + 1):
            cap = rc * self.v[i] * R * R / dt
            gw = k * self.f[i - 1] * R / dr if i > 0 else 0.0
            ge = k * self.f[i] * R / dr if i < N else 0.0
            a[i], c[i] = -gw, -ge
            b[i] = cap + gw + ge
            d[i] = cap * T[i] + S * self.v[i] * R * R
        b[N] += heff * R
        d[N] += heff * R * T_AMB - q_die * R
        for i in range(1, N + 1):
            w = a[i] / b[i - 1]
            b[i] -= w * c[i - 1]
            d[i] -= w * d[i - 1]
        T[N] = d[N] / b[N]
        for i in range(N - 1, -1, -1):
            T[i] = (d[i] - c[i] * T[i + 1]) / b[i]


# ---------------------------------------------------------------- Geometrie des Stichplans
def plan_geometry(block, passes):
    """block: dict(shape, d, L, A) [mm, mm²]; passes: Liste dict(shape, mode, dim, len).
    Ergänzt je Stich d, L, A, ratio, phi, phi_cum, h_prev, h_new, dh, ok, msg."""
    V = block["A"] * block["L"]
    prev = dict(shape=block["shape"], d=block["d"], L=block["L"], A=block["A"])
    phi_cum = 0.0
    out = []
    for p in passes:
        r = dict(p)
        r["ok"], r["msg"] = False, ""
        try:
            if p["mode"] == "dim":
                d = float(p["dim"])
                if not d > 0:
                    raise ValueError
                A = area(p["shape"], d)
                L = V / A
            else:
                L = float(p["len"])
                if not L > 0:
                    raise ValueError
                A = V / L
                d = dim_from_area(p["shape"], A)
        except (TypeError, ValueError):
            r["msg"] = "Maß oder Länge eingeben."
            out.append(r)
            prev = None
            break
        if prev is None:
            break
        ratio = prev["A"] / A
        r.update(d=d, L=L, A=A, ratio=ratio, phi=math.log(ratio))
        hp = height_before(prev["shape"], prev["d"], p["shape"])
        r.update(h_prev=hp, h_new=d, dh=hp - d, prev=prev)
        if ratio <= 1.0 + 1e-6:
            r["msg"] = "Der Querschnitt nimmt nicht ab. Im Stichplan ist nur Recken vorgesehen."
            out.append(r)
            break
        phi_cum += r["phi"]
        r["phi_cum"] = phi_cum
        r["ok"] = True
        out.append(r)
        prev = dict(shape=p["shape"], d=d, L=L, A=A)
    return out


# ---------------------------------------------------------------- Temperatur und Kraft je Stich
def simulate_plan(block, geo, mat, kf_model, proc, N=30):
    """
    block: dict(shape, d, L, A) in mm
    geo:   Ergebnis von plan_geometry (nur Stiche mit ok=True werden gerechnet)
    mat:   dict(rho, cp, k)
    kf_model: dict(type='st', alpha, n, Q, lnA) oder dict(type='hs', A, m1..m9)
    proc:  T_start, t_transfer, v [mm/s], t_dwell, t_aux, t_pause, sB_rel, b_die [mm], T_die, h_c,
           lam_die, h_air, eps, eta_F, mu
    """
    rho, cp, k = mat["rho"], mat["cp"], mat["k"]
    rc = rho * cp
    e1 = math.sqrt(k * rc)
    e2 = math.sqrt(proc["lam_die"] * RC_DIE)
    sol = Radial(N, proc["T_start"])
    t = 0.0
    hist_t, hist_c, hist_s, hist_m = [0.0], [sol.T[0]], [sol.T[N]], [sol.mean()]
    contacts = []
    bands = []

    def run(duration, R, h_air_frac, q_die, S, dt_max):
        nonlocal t
        if duration <= 0:
            return
        n = max(1, int(math.ceil(duration / dt_max)))
        dt = duration / n
        for _ in range(n):
            sol.step(R, k, rc, dt, proc["h_air"], h_air_frac, q_die, S, proc["eps"])
            t += dt
            hist_t.append(t)
            hist_c.append(sol.T[0])
            hist_s.append(sol.T[N])
            hist_m.append(sol.mean())

    def r_eq(shape_a, d_a, shape_b, d_b, w):
        A = (1 - w) * area(shape_a, d_a) + w * area(shape_b, d_b)
        P = (1 - w) * perimeter(shape_a, d_a) + w * perimeter(shape_b, d_b)
        return 2.0 * A / P * 1e-3, A, P

    # Transport Ofen -> Presse
    R0 = 2.0 * block["A"] / perimeter(block["shape"], block["d"]) * 1e-3
    run(proc["t_transfer"], R0, 1.0, 0.0, 0.0, 0.5)
    start = (sol.T[0], sol.T[N], sol.mean())
    rows = []
    n_ok = sum(1 for g in geo if g.get("ok"))
    v = proc["v"]
    for g in geo:
        if not g.get("ok"):
            break
        prev = g["prev"]
        n_rot = int(g.get("nrot") or DEFAULT_NROT[g["shape"]])
        n_rot = max(1, n_rot)
        dh = max(g["dh"], 0.02 * g["h_prev"])
        n_b, sB, s_last, feed, _s_max = bite_layout(prev["L"], g["h_prev"], g["ratio"], proc)
        bc_auto = contact_width_auto(g["shape"], g["d"], g["h_prev"])
        bc = g.get("bk") or bc_auto
        t_c = dh / v + proc["t_dwell"]
        t_cycle = t_c + proc["t_aux"]
        # Bezugsscheibe in Stabmitte liegt im Biss k_mid; vorher und nachher werden die übrigen Bisse geschmiedet
        k_mid = min(n_b - 1, int((prev["L"] / 2.0) // sB))
        h_mean = 0.5 * (g["h_prev"] + g["h_new"])
        phid = v / h_mean
        dphi_k = g["phi"] / n_rot
        phi0 = g["phi_cum"] - g["phi"]
        t0 = t
        kf_list, F_list, Tc_list = [], [], []
        dT_def_sum, dT_die_sum, dT_rest_sum = 0.0, 0.0, 0.0
        T_mean_start = sol.mean()
        R_0, _, _ = r_eq(prev["shape"], prev["d"], g["shape"], g["d"], 0.0)
        run(k_mid * n_rot * t_cycle, R_0, 1.0, 0.0, 0.0, 0.5)            # Bisse vor der Bezugsscheibe
        for kk in range(n_rot):
            w0, w1 = kk / n_rot, (kk + 1) / n_rot
            R_a, A_a, P_a = r_eq(prev["shape"], prev["d"], g["shape"], g["d"], w0)
            R_b, A_b, P_b = r_eq(prev["shape"], prev["d"], g["shape"], g["d"], w1)
            t_aux = proc["t_aux"]
            if t_aux > 0:                                                   # Nebenzeit auf dem Untersattel
                _Ts_r, E_r = contact_two_halfspaces(sol.T[N], proc["T_die"], proc.get("h_rest", 0.0), e1, e2, t_aux)
                f_r = min(1.0, bc / P_a)
                dT_rest_sum += f_r * E_r * P_a * 1e-3 / (A_a * 1e-6 * rc)
                run(t_aux, R_a, 1.0 - f_r, f_r * E_r / t_aux, 0.0, max(t_aux / 20.0, 0.05))
            # Fließspannung bei aktueller Mitteltemperatur (+ halbe Umformerwärmung)
            Tm = sol.mean()
            p0, p1 = phi0 + kk * dphi_k, phi0 + (kk + 1) * dphi_k
            kf = kf_mean(kf_model, Tm, p0, p1, phid)
            dT_def = BETA_TQ * kf * 1e6 * dphi_k / (proc["eta_F"] * rc)
            kf = kf_mean(kf_model, Tm + 0.5 * dT_def, p0, p1, phid)
            dT_def = BETA_TQ * kf * 1e6 * dphi_k / (proc["eta_F"] * rc)
            # Werkzeugkontakt (lokal, geschlossene Lösung)
            Ts_contact, E = contact_two_halfspaces(sol.T[N], proc["T_die"], proc["h_c"], e1, e2, t_c)
            f_die = min(1.0, 2.0 * bc / P_b)
            q_die = f_die * E / t_c
            dT_die_sum += f_die * E * P_b * 1e-3 / (A_b * 1e-6 * rc)
            dT_def_sum += dT_def
            run(t_c, R_b, 1.0 - f_die, q_die, rc * dT_def / t_c, max(t_c / 20.0, 0.02))
            contacts.append((t, Ts_contact))
            Tc_list.append(Ts_contact)
            # Presskraft am Hubende
            h_end = g["h_new"]
            Q = 1.155 * max(0.75 + 0.25 * h_end / sB, 1.0 + proc["mu"] * sB / (4.0 * h_end))
            F = kf * Q * sB * bc * 1e-6           # MPa · mm² = N -> MN
            kf_list.append(kf)
            F_list.append(F)
        R_1, _, _ = r_eq(prev["shape"], prev["d"], g["shape"], g["d"], 1.0)
        run((n_b - k_mid - 1) * n_rot * t_cycle, R_1, 1.0, 0.0, 0.0, 0.5)  # Bisse nach der Bezugsscheibe
        t1 = t
        rows.append(dict(
            n_rot=n_rot, n_b=n_b, strokes=n_rot * n_b, sB=sB, s_last=s_last, feed=feed, rot_step=180.0 / n_rot,
            bc=bc, bc_auto=bc_auto, A_K=sB * bc,
            t_c=t_c, dh=dh, phid=phid, kf=max(kf_list), F=max(F_list), t_pass=t1 - t0, t_end=t1,
            T_core=sol.T[0], T_surf=sol.T[N], T_mean=sol.mean(), T_contact=min(Tc_list),
            dT_def=dT_def_sum, dT_die=dT_die_sum + dT_rest_sum, dT_rest=dT_rest_sum,
            dT_mean=sol.mean() - T_mean_start,
        ))
        bands.append((t0, t1))
        if len(rows) < n_ok:
            R_end, _, _ = r_eq(prev["shape"], prev["d"], g["shape"], g["d"], 1.0)
            run(proc["t_pause"], R_end, 1.0, 0.0, 0.0, 0.5)
    return dict(rows=rows, t=hist_t, core=hist_c, surf=hist_s, mean=hist_m, contacts=contacts, bands=bands,
                t_total=t, start=start)
