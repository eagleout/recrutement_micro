# -*- coding: utf-8 -*-
"""Génère app/index.html à partir de app/template.html et du contenu métier (excel/content.py).

Usage : python app/gen.py
Les candidats d'exemple sont repris de la liste de démo de excel/build.py.
"""
import json, sys, datetime as dt
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "excel"))
import content as C  # noqa: E402

src = (ROOT / "excel" / "build.py").read_text(encoding="utf-8")
i0 = src.index("demo = [\n")
block = src[i0:src.index("if not DEMO:", i0)]
ns = {"dt": dt, "D": lambda m, d: dt.date(2026, m, d), "C1": 0, "C2": 1, "FULL_OK": ("Oui", "Reçue", "OK", "OK")}
exec(block, ns)

iso = lambda d: d.isoformat() if d else ""
examples = []
for i, r in enumerate(ns["demo"]):
    (nom, pre, cr, po, source, dcand, dip, exp, dispo, pret, et, stat, dla, nxt, ntel,
     conf, demb, comm, notes, evr, forts, vig) = r
    c = {"id": f"ex{i}", "_ex": True, "nom": nom, "prenom": pre, "tel": f"06 00 00 {i // 10:02d} {i % 100:02d}", "email": "",
         "creche": cr, "poste": po, "source": source, "dateCand": iso(dcand), "diplome": dip, "exp": exp, "dispo": iso(dispo),
         "pret": pret, "etape": et, "statut": stat, "derniereAction": iso(dla), "prochaine": nxt,
         "conf": dict(zip(("diplome", "honor", "vacc", "refs"), conf)), "dateEmb": iso(demb), "comment": comm,
         "eval": {"notes": list(notes) if notes else [None] * 12, "evaluateur": evr,
                  "date": iso(dla) if notes and et >= 4 else "", "forts": forts, "vigilance": vig}}
    if ntel:
        c["noteTel"] = ntel
    examples.append(c)

keys = ("BLOCS", "POIDS", "POSTES", "ETAPES", "STATUTS", "SOURCES", "DIPLOMES", "CONTRATS", "OUINON", "HONOR", "VACCINS",
        "REFS", "ECHELLE", "PREREQUIS", "FICHES", "FICHE_SECTIONS", "QUESTIONS", "SCRIPT_TEL", "MODELES", "ONBOARDING",
        "REGLEMENTAIRE")
data = {k: getattr(C, k) for k in keys}
data["EXAMPLES"] = examples
js = json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")
html = (ROOT / "app" / "template.html").read_text(encoding="utf-8").replace("__DATA__", js)
(ROOT / "app" / "index.html").write_text(html, encoding="utf-8")
print(f"app/index.html : {len(html)} octets, {len(examples)} candidats d'exemple")
