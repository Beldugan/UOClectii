#!/usr/bin/env python3
"""Generează pachetele PDF pentru fiecare oră UOC (FIMIM) din Google Calendar
și datele pentru site-ul beldugan.github.io/UOClectii."""
import hashlib, json, os, re, shutil, subprocess, tempfile
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(TTFont("DV", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

SRC = "/mnt/user-data/uploads/2026 2027"
OUT = "/home/claude/uoclectii"
EV = "/tmp/claude-0/uoc/ev1.json"

F = dict(
    MF=f"{SRC}/AR IF/Mecanica Fluidelor/MF_Indrumar_laborator_2026-2027.pdf",
    SAA=f"{SRC}/AR IF/Sisteme auxiliare ale autovehiculelor/SAA_Indrumar_laborator_2026-2027.pdf",
    DA=f"{SRC}/AR IF/Diagnosticarea autovehiculelor/Diagnosticarea_autovehiculelor_Indrumar_laborator_2026-2027.pdf",
    MAM=f"{SRC}/AR IF/Mecatronica automobilului modern I/MAM_I_Indrumar_laborator_2026-2027.pdf",
    PC=f"{SRC}/AR IF/PCMAI proiect/PCMAI_Indrumar_proiect_2026-2027.pdf",
    IAII=f"{SRC}/IAII/IAII_Laboratoare_Beldugan_2026-2027.pdf",
)
XLSX = f"{SRC}/AR IF/PCMAI proiect/PCMAI_Calcul_termic_2026-2027.xlsx"
if not os.path.exists(XLSX):
    XLSX = "/mnt/user-data/outputs/PCMAI_Calcul_termic_2026-2027.xlsx"

MAN = dict(MF="Îndrumar de laborator – Mecanica fluidelor", SAA="Îndrumar de laborator – Sisteme auxiliare ale autovehiculelor",
           DA="Îndrumar de laborator – Diagnosticarea autovehiculelor", MAM="Îndrumar de laborator – Mecatronica automobilului modern I",
           PC="Îndrumar de proiect – PCMAI", IAII="Programa propusă a laboratoarelor IAII")

# ------------------------------------------------ conținut pe lucrări: (titlu, [(doc, a, b), ...])
MF = [
    ("Vâscozitatea uleiurilor de motor în funcție de temperatură", 6, 10),
    ("Compresibilitatea lichidului de frână și efectul aerului din circuit", 11, 15),
    ("Legea lui Pascal: transmiterea hidraulică a forței în sistemul de frânare", 16, 19),
    ("Ecuația lui Bernoulli și măsurarea debitului: tubul Venturi și tubul Pitot", 20, 23),
    ("Regimuri de curgere. Experiența lui Reynolds", 24, 27),
    ("Pierderi de sarcină longitudinale și locale într-un circuit de răcire", 28, 32),
    ("Definitivarea și susținerea referatelor", 33, 38),
]
SAA = [
    ("Sistemul de injecție de benzină multipunct – caracteristicile de funcționare pe standul motorului 1.4 MPI", 5, 9),
    ("Modelarea duratei de injecție pe baza debitului de aer – aplicație virtuală LabVIEW", 10, 12),
    ("Sistemul de injecție common-rail – reglarea presiunii în rampă și injecția multiplă (1.5 dCi)", 13, 16),
    ("Sistemul de ungere – caracteristica presiunii uleiului", 17, 19),
    ("Sistemul de răcire și managementul termic", 20, 23),
    ("Sistemul de supraalimentare – presiunea de supraalimentare și actuatorii turbosuflantei", 24, 27),
    ("Sistemul de pornire și bilanțul energetic al sistemelor stop-start și mild-hybrid 48 V", 28, 32),
    ("Sistemul de climatizare și pompa de căldură – determinarea parametrilor de funcționare", 33, 37),
]
DA = [
    ("Diagnoza computerizată și rețeaua de date CAN", 6, 11),
    ("Senzori și actuatori – date live și semnale reale", 12, 17),
    ("Starea mecanică a motorului", 18, 22),
    ("Sistemul de alimentare MAS și analiza gazelor de evacuare", 23, 27),
    ("Motorul diesel common rail și sistemele de post-tratare (EGR, DPF, SCR)", 28, 33),
    ("Sistemul de frânare cu ABS/ESC și sistemele ADAS", 34, 39),
    ("Vehicule electrice și hibride; diagnoza pe bază de date", 40, 48),
]
MAM = [
    ("Microcontrolere: intrări/ieșiri, convertor A/D, timere, întreruperi și PWM", 6, 9),
    ("Condiționarea semnalelor și conversia analog-digitală", 10, 13),
    ("Senzorul pedalei de accelerație cu două piste: plauzibilitate și reacția la defect", 14, 16),
    ("Senzori de temperatură: liniarizare și constanta de timp", 17, 20),
    ("Senzorul de presiune absolută din colector (MAP): calibrare și calculul densității aerului", 21, 23),
    ("Senzori de turație și poziție: decodarea roții 60−2 prin captură de intrare", 24, 27),
    ("Senzori ultrasonici de parcare: timpul de propagare și compensarea cu temperatura", 28, 30),
    ("Senzori inerțiali (IMU): unghiurile de ruliu și tangaj, filtrul complementar", 31, 33),
    ("Clapeta electronică de accelerație: punte H, comandă PWM și regulator PID", 34, 37),
    ("Magistrala CAN: construirea mesajelor, arbitrajul și decodarea semnalelor", 38, 44),
]
PCS = [  # etape de proiect: (titlu, ore, pagini)
    ("Prezentarea proiectului și a temei; alegerea parametrilor inițiali; motorul de referință", 2, [(6, 6), (17, 17), (23, 23), (34, 35)]),
    ("Calculul procesului de admisie", 3, [(7, 8), (17, 17), (23, 23)]),
    ("Calculul procesului de comprimare", 2, [(9, 9), (18, 18), (24, 24)]),
    ("Arderea: MAS – izocoră; MAC – întârzierea la autoaprindere și arderea izocoră", 3, [(10, 11), (18, 18), (24, 24)]),
    ("Arderea izobară la MAC; rapoartele λp și ρ", 2, [(10, 11), (24, 25)]),
    ("Destinderea; verificarea temperaturii gazelor reziduale", 2, [(12, 12), (19, 19), (25, 25)]),
    ("Parametrii indicați și efectivi", 3, [(13, 13), (19, 19), (26, 26)]),
    ("Parametrii constructivi; compararea cu motorul de referință", 3, [(14, 14), (19, 19), (22, 22), (26, 26), (28, 28)]),
    ("Diagrama indicată p–V și p–α", 3, [(15, 15), (20, 20), (26, 26)]),
    ("Caracteristica exterioară de turație; definitivarea memoriului", 3, [(16, 16), (21, 21), (27, 28), (30, 31)]),
    ("Colocviu – susținerea proiectului", 2, [(30, 31)]),
]
ST = ["Norme de securitate a muncii. Prezentarea laboratorului, a aparaturii și a platformei de achiziție",
      "Caracteristicile statice ale traductoarelor: sensibilitate, liniaritate, histerezis",
      "Traductoare rezistive de temperatură (Pt100/Pt1000) – condiționare în punte, liniarizare",
      "Termistoare NTC/PTC și termocupluri – condiționare, compensarea joncțiunii reci",
      "Traductoare de deplasare liniară și unghiulară: potențiometric, LVDT",
      "Encodere incrementale/absolute – decodarea semnalelor A/B, măsurarea turației",
      "Traductoare tensometrice – măsurarea forței și a deformației (punte Wheatstone, celulă de sarcină)",
      "Traductoare de presiune piezorezistive – ieșire analogică și digitală",
      "Senzori inductivi și capacitivi de proximitate (industriali, PNP/NPN)",
      "Senzori optici și ultrasonici de distanță/prezență",
      "Accelerometre și IMU – măsurarea vibrațiilor",
      "Interfațarea industrială: bucla de curent 4–20 mA, conversia analog-numerică",
      "Senzori digitali și rețele de senzori: I²C, SPI, Modbus/IO-Link – achiziția și prelucrarea datelor",
      "Aplicație finală de achiziție de date. Verificarea referatelor – colocviu de laborator"]
API = ["Norme de securitate a muncii. Mediul de lucru Python + OpenCV; citirea și afișarea imaginilor",
       "Reprezentarea imaginilor digitale; spații de culoare și conversii (RGB, HSV, niveluri de gri)",
       "Histograma; operații punctuale – ajustarea contrastului, egalizarea histogramei",
       "Binarizarea imaginilor: prag global, metoda Otsu, prag adaptiv",
       "Filtrarea în domeniul spațial; reducerea zgomotului (filtre medie, Gauss, median)",
       "Detecția muchiilor: operatori Sobel, Prewitt, Canny",
       "Operații morfologice pe imagini binare: eroziune, dilatare, deschidere, închidere",
       "Etichetarea componentelor conexe; proprietăți geometrice ale obiectelor",
       "Transformări geometrice; calibrarea camerei și măsurarea dimensională din imagini",
       "Transformata Hough – detecția liniilor și a cercurilor",
       "Filtrarea în domeniul frecvență (transformata Fourier)",
       "Detecția și potrivirea trăsăturilor: template matching, ORB/SIFT",
       "Aplicație industrială: inspecția vizuală a pieselor și detecția defectelor",
       "Verificarea finală a lucrărilor – colocviu de laborator"]


def lab_units(doc, labs, groups, intro=True):
    """groups: listă de liste de indici (0-based) de lucrări pentru fiecare ședință."""
    out = []
    for k, g in enumerate(groups):
        pages, parts, titles = [], [], []
        if k == 0 and intro:
            pages.append((doc, 3, labs[0][1] - 1)); parts.append(f"{MAN[doc]} – Introducere și norme de securitate")
        for i, suf in g:
            t, a, b = labs[i]
            titles.append(f"Lucrarea {i + 1}{suf}. {t}")
            pages.append((doc, a, b)); parts.append(f"{MAN[doc]} – Lucrarea {i + 1}{suf}")
        out.append(dict(title=" · ".join(titles), pages=pages, parts=parts, nr=", ".join(f"L{i+1}{s}" for i, s in g)))
    return out


def pc_units(n):
    tot = sum(h for _, h, _ in PCS)
    bounds, c = [], 0
    for t, h, p in PCS:
        bounds.append((c, c + h)); c += h
    out = []
    for k in range(n):
        lo, hi = k * tot / n, (k + 1) * tot / n
        st = [i for i, (a, b) in enumerate(bounds) if a < hi - 1e-9 and b > lo + 1e-9]
        seen, pages = set(), []
        for i in st:
            for a, b in PCS[i][2]:
                if (a, b) not in seen:
                    seen.add((a, b)); pages.append(("PC", a, b))
        pages.sort(key=lambda x: x[1])
        titles = [f"Etapa {i + 1}. {PCS[i][0]}" for i in st]
        parts = [f"{MAN['PC']} – " + ", ".join(f"p. {a}–{b}" if a != b else f"p. {a}" for _, a, b in pages),
                 "Foaia de calcul PCMAI (Excel) – de descărcat separat"]
        out.append(dict(title=" · ".join(titles), pages=pages, parts=parts, nr=", ".join(f"E{i+1}" for i in st),
                        extra=[["Foaia de calcul termic (.xlsx)", "files/PCMAI_Calcul_termic_2026-2027.xlsx"]]))
    return out


def iaii_units(names, page, groups):
    out = []
    for g in groups:
        out.append(dict(title=" · ".join(f"Lucrarea {i + 1}. {names[i]}" for i in g),
                        pages=[("IAII", page, page)],
                        parts=["Programa propusă a laboratoarelor (de validat cu titularul de curs)",
                               "Îndrumarul de laborator este în pregătire – se adaugă aici pe măsură ce este redactat"],
                        nr=", ".join(f"L{i+1}" for i in g)))
    return out


S = lambda *ix: [(i, "") for i in ix]
UNITS = {
    "mf": lab_units("MF", MF, [S(i) for i in range(7)]),
    "saa": lab_units("SAA", SAA, [S(0), S(1, 3), S(2), S(4), S(5), S(6), S(7)]),
    "da": lab_units("DA", DA, [S(i) for i in range(7)]),
    "mam": lab_units("MAM", MAM, [[(0, "a")], [(0, "b")], S(1), S(2), S(3), S(4), S(5), S(6),
                                   [(7, "a")], [(7, "b")], [(8, "a")], [(8, "b")], S(9)]),
    "st": iaii_units(ST, 1, [[i] for i in range(12)] + [[12, 13]]),
    "api": iaii_units(API, 2, [[i] for i in range(14)]),
}

# ------------------------------------------------ clasificarea evenimentelor
PROG = {"AR II": "ar2", "AR III": "ar3", "AR IFR": "arifr", "AR IV": "ar4", "IAII III": "iaii3"}
PROG_NAME = {"ar2": "Autovehicule rutiere – anul II", "ar3": "Autovehicule rutiere – anul III",
             "arifr": "Autovehicule rutiere IFR – anul III", "ar4": "Autovehicule rutiere – anul IV",
             "iaii3": "Informatică aplicată în ingineria industrială – anul III"}
DISC = [("Mecanica fluidelor", "mf", "Mecanica fluidelor", "laborator"),
        ("PCMAI", "pc", "Procese și caracteristici ale motoarelor cu ardere internă", "proiect"),
        ("Diagnosticarea", "da", "Diagnosticarea autovehiculelor", "laborator"),
        ("Mecatronica", "mam", "Mecatronica automobilului modern I", "laborator"),
        ("Sisteme auxiliare", "saa", "Sisteme auxiliare ale autovehiculelor", "laborator"),
        ("Senzori", "st", "Senzori și traductoare", "laborator"),
        ("Analiza", "api", "Analiza și prelucrarea imaginilor", "laborator")]
ZI = ["luni", "marți", "miercuri", "joi", "vineri", "sâmbătă", "duminică"]


def cover(path, e):
    W, H = A4
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(e["title"])
    c.setFillColorRGB(0.07, 0.25, 0.45); c.rect(0, H - 110, W, 110, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1); c.setFont("DV", 9.5)
    c.drawString(50, H - 40, "UNIVERSITATEA „OVIDIUS” DIN CONSTANȚA · FIMIM · An universitar 2026–2027")
    c.setFont("DVB", 19)
    c.drawString(50, H - 75, f"{e['prog']}" + (f" · {e['sgr']}" if e["sgr"] else "") + f" · S{e['week']}")
    c.setFillColorRGB(0.1, 0.1, 0.1)
    y = H - 150
    y = wrap(c, f"{e['module']} – {e['kind']}", 50, y, W - 100, "DV", 12, 16)
    y = wrap(c, f"{ZI[e['wd']]} {e['date'][8:10]}.{e['date'][5:7]}.{e['date'][:4]}, {e['time']}–{e['end']} · {e['loc']}",
             50, y, W - 100, "DV", 11, 15)
    y -= 12
    y = wrap(c, e["title"], 50, y, W - 100, "DVB", 15, 20)
    y -= 18
    c.setFont("DVB", 11); c.drawString(50, y, "Conținutul pachetului"); y -= 18
    pg = 2
    for doc, a, b in e["_pages"]:
        n = b - a + 1
        rng = f"p. {pg}–{pg + n - 1}" if n > 1 else f"p. {pg}"
        src = f"p. {a}–{b}" if n > 1 else f"p. {a}"
        y = wrap(c, f"{rng}   {MAN[doc]}   (sursa: {src})", 60, y, W - 120, "DV", 10.5, 14); y -= 3
        pg += n
    for lbl, _ in e.get("extra", []):
        y -= 6; y = wrap(c, f"Fișier separat pe pagina orei: {lbl}", 60, y, W - 120, "DV", 10.5, 14)
    c.setFont("DV", 8.5); c.setFillColorRGB(0.4, 0.4, 0.4)
    c.drawString(50, 40, "Drd. ing. Adrian-Mircea Beldugan · beldugan.github.io/UOClectii")
    c.showPage(); c.save()


def wrap(c, text, x, y, w, font, size, lead):
    c.setFont(font, size)
    words, line = text.split(), ""
    for wd in words:
        t = (line + " " + wd).strip()
        if pdfmetrics.stringWidth(t, font, size) > w and line:
            c.drawString(x, y, line); y -= lead; line = wd
        else:
            line = t
    if line:
        c.drawString(x, y, line); y -= lead
    return y


def main():
    from datetime import date, timedelta
    ev = json.load(open(EV))["events"]
    ev.sort(key=lambda e: e["start"]["dateTime"])
    mondays = sorted({(date.fromisoformat(e["start"]["dateTime"][:10]) -
                       timedelta(days=date.fromisoformat(e["start"]["dateTime"][:10]).weekday())) for e in ev})
    wk = {m: i + 1 for i, m in enumerate(mondays)}
    series = {}
    for e in ev:
        s = e["summary"]
        m = re.match(r"\[([^\]]+)\]", s); prog = PROG[m.group(1)]
        d = next(x for x in DISC if x[0] in s)
        sg = re.search(r"sgr (\d)", s)
        key = (prog, d[1], sg.group(1) if sg else "")
        series.setdefault(key, []).append((e, d))
    if os.path.exists(OUT + "/pdf"): shutil.rmtree(OUT + "/pdf")
    os.makedirs(OUT + "/pdf", exist_ok=True); os.makedirs(OUT + "/data", exist_ok=True)
    os.makedirs(OUT + "/files", exist_ok=True)
    shutil.copy(XLSX, OUT + "/files/PCMAI_Calcul_termic_2026-2027.xlsx")
    ore, calmap = [], []
    tmp = tempfile.mkdtemp()
    for (prog, dk, sg), lst in series.items():
        units = pc_units(len(lst)) if dk == "pc" else UNITS[dk]
        assert len(units) == len(lst), (prog, dk, sg, len(units), len(lst))
        for k, ((e, d), u) in enumerate(zip(lst, units)):
            st = e["start"]["dateTime"]; dt = date.fromisoformat(st[:10])
            oid = f"{prog}-{dk}{('-s' + sg) if sg else ''}-{st[:10]}-{st[11:13]}{st[14:16]}"
            o = dict(id=oid, cls=prog, disc=dk, type="p" if d[3] == "proiect" else "l",
                     title=u["title"], nr=u["nr"], date=st[:10], time=st[11:16], end=e["end"]["dateTime"][11:16],
                     week=wk[dt - timedelta(days=dt.weekday())], module=d[2],
                     sgr=f"subgrupa {sg}" if sg else "", loc=e.get("location", "").replace(", FIMIM UOC", ""),
                     seq=k + 1, n=len(lst), parts=u["parts"])
            if u.get("extra"): o["extra"] = u["extra"]
            h = hashlib.sha1((oid + json.dumps(u["pages"])).encode()).hexdigest()[:6]
            o["pdf"] = f"pdf/{oid}-{h}.pdf"
            meta = dict(o, prog=PROG_NAME[prog], kind=d[3], wd=dt.weekday(), _pages=u["pages"])
            cv = os.path.join(tmp, "c.pdf"); cover(cv, meta)
            args = ["qpdf", "--empty", "--pages", cv, "1"]
            for doc, a, b in u["pages"]:
                args += [F[doc], f"{a}-{b}"]
            args += ["--", "--object-streams=generate", "--compress-streams=y", "--recompress-flate",
                     "--remove-unreferenced-resources=yes", OUT + "/" + o["pdf"]]
            r = subprocess.run(args, capture_output=True, text=True)
            if r.returncode not in (0, 3): raise RuntimeError(r.stderr)
            ore.append(o)
            calmap.append(dict(event_id=e["id"], summary=e["summary"], start=st, id=oid, title=u["title"],
                               nr=u["nr"], week=o["week"], description=e.get("description", "")))
    ore.sort(key=lambda o: (o["date"], o["time"]))
    json.dump(ore, open(OUT + "/data/ore.json", "w"), ensure_ascii=False, indent=0)
    json.dump(calmap, open("/tmp/claude-0/uoc/calmap.json", "w"), ensure_ascii=False, indent=1)
    print(len(ore), "ore")


if __name__ == "__main__":
    main()
