# -*- coding: utf-8 -*-
"""
verify_all.py — Vérification complète automatique de tous les commerciaux
=========================================================================
Détecte automatiquement pour CHAQUE commercial :
  1. Écart tonnage vs objectif théorique
  2. Vrais doublons (même agriculteur + même usine + même tonnage exact)
  3. Multi-parcelles légitimes (même agriculteur + usines/tonnages différents)
  4. Agriculteurs avec tonnage anormal (0t, négatif, ou aberrant)
  5. Agriculteurs sans région ou sans zone
  6. Surplus/manque par rapport à l'objectif

Usage: python verify_all.py
       python verify_all.py --fix     (supprime les vrais doublons)
       python verify_all.py --export  (exporte rapport en CSV)
"""
import sys, io, argparse
from collections import defaultdict, Counter
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from supabase import create_client

# ── Config ────────────────────────────────────────────────────
SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

# Tonnages officiels attendus par commercial
EXPECTED = {
    "FEDI":             31_690,
    "MAKKI BEN SALAH":  21_365,
    "KHALIL":           15_705,
    "ACHREF AJLANI":    14_140,
    "JILANI OBAY":       7_410,
}
TOTAL_EXPECTED = sum(EXPECTED.values())  # 90,310t

# Seuils d'alerte
MAX_TONNAGE_PER_FARMER  = 10_000   # au-delà = suspect
MIN_TONNAGE_PER_FARMER  = 10       # en-dessous = suspect
TOLERANCE_PCT           = 0.5      # ±0.5% tolérance sur total commercial

# ── Helpers ──────────────────────────────────────────────────
def fetch_all(sb, table, columns="*"):
    """Lit TOUTES les lignes avec pagination Supabase."""
    all_rows, offset, page = [], 0, 1000
    while True:
        batch = sb.table(table).select(columns).range(offset, offset+page-1).execute().data
        if not batch: break
        all_rows.extend(batch)
        if len(batch) < page: break
        offset += page
    return all_rows

def color(txt, code): return f"\033[{code}m{txt}\033[0m"
def green(t):  return color(t, "92")
def red(t):    return color(t, "91")
def yellow(t): return color(t, "93")
def cyan(t):   return color(t, "96")
def bold(t):   return color(t, "1")

# ── Main ─────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("--fix",    action="store_true", help="Supprimer les vrais doublons")
parser.add_argument("--export", action="store_true", help="Exporter rapport CSV")
args = parser.parse_args()

print(bold("="*65))
print(bold(f"  VÉRIFICATION COMPLÈTE — TOUS LES COMMERCIAUX"))
print(bold(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"))
print(bold("="*65))

sb = create_client(SUPABASE_URL, SUPABASE_KEY)
print("\n  Connexion Supabase...", end="")
data = fetch_all(sb, "agriculteurs", "id,commercial,nom,usine,zone,region,tonnage_total,accessibilite")
print(f" OK — {len(data)} lignes chargées\n")

# ── Organiser par commercial ──────────────────────────────────
by_comm = defaultdict(list)
for r in data:
    comm = r.get("commercial", "INCONNU") or "INCONNU"
    by_comm[comm].append(r)

# ── Rapport global ────────────────────────────────────────────
all_problems   = []   # pour export CSV
total_reel     = 0
total_doublons = 0
total_fixes    = 0

for comm in sorted(EXPECTED.keys()):
    rows     = by_comm.get(comm, [])
    expected = EXPECTED[comm]
    actual   = sum(float(r.get("tonnage_total", 0) or 0) for r in rows)
    ecart    = actual - expected
    ecart_pct = (ecart / expected * 100) if expected > 0 else 0
    total_reel += actual

    problems = []
    fixes_needed = []

    print(bold("─"*65))
    print(bold(f"  👤 {comm}"))
    print(f"     Lignes   : {len(rows)} | Tonnage réel: {actual:,.0f}t | Attendu: {expected:,.0f}t")

    # ── 1. Écart total ──────────────────────────────────────
    if abs(ecart_pct) <= TOLERANCE_PCT:
        print(f"     Écart    : {green(f'{ecart:+,.0f}t ({ecart_pct:+.2f}%) ✅')}")
    elif ecart > 0:
        msg = f"{ecart:+,.0f}t ({ecart_pct:+.1f}%) — SURPLUS ⚠️"
        print(f"     Écart    : {yellow(msg)}")
        problems.append(("SURPLUS", comm, "TOTAL", "-", f"{ecart:+,.0f}t en trop"))
    else:
        msg = f"{ecart:+,.0f}t ({ecart_pct:+.1f}%) — MANQUE ❌"
        print(f"     Écart    : {red(msg)}")
        problems.append(("MANQUE", comm, "TOTAL", "-", f"{ecart:,.0f}t manquants"))

    if not rows:
        print(f"     {red('❌ AUCUN AGRICULTEUR trouvé pour ce commercial !')}")
        problems.append(("VIDE", comm, "-", "-", "Aucun agriculteur en base"))
        all_problems.extend(problems)
        continue

    # ── 2. Vrais doublons (même tonnage exact) ──────────────
    groups = defaultdict(list)
    for r in rows:
        key = (r.get("nom",""), r.get("usine","") or "")
        groups[key].append({"id": r.get("id"), "ton": float(r.get("tonnage_total",0) or 0)})

    vrais_doublons = []
    for (nom, usine), entries in groups.items():
        if len(entries) < 2: continue
        ton_cnt = Counter(e["ton"] for e in entries)
        for ton, cnt in ton_cnt.items():
            if cnt >= 2:
                ids = [e["id"] for e in entries if e["ton"] == ton]
                vrais_doublons.append({
                    "nom": nom, "usine": usine, "tonnage": ton,
                    "count": cnt, "ids": ids,
                    "supprimer_ids": ids[1:]  # garder le premier, supprimer le reste
                })

    if vrais_doublons:
        print(f"     {red(f'❌ {len(vrais_doublons)} VRAI(S) DOUBLON(S) détecté(s):')}")
        for d in vrais_doublons:
            total_surplus = d["tonnage"] * (d["count"] - 1)
            print(f"        → {d['nom']} → {d['usine']}: "
                  f"{d['tonnage']:,.0f}t × {d['count']} fois = "
                  f"{red(f'+{total_surplus:,.0f}t en trop')}")
            problems.append(("DOUBLON", comm, d["nom"], d["usine"],
                             f"{d['tonnage']:,.0f}t × {d['count']} = +{total_surplus:,.0f}t"))
            fixes_needed.extend(d["supprimer_ids"])

        # ⛔ JAMAIS de suppression automatique même avec --fix
        # Raison: tonnages identiques peuvent être 2 parcelles légitimes
        # (ex: KHALED CHATER a 2 champs de 350t chacun → même tonnage, pas un doublon)
        # Solution: le directeur corrige manuellement via dashboard "Gestion Agriculteurs"
        print(f"        ℹ️  Action requise: vérifier dans le dashboard")
        print(f"        → Gestion Agriculteurs → chercher l'agriculteur")
        print(f"        → Si 2 parcelles réelles: garder les 2")
        print(f"        → Si upload en double: supprimer manuellement")
    else:
        print(f"     Doublons : {green('✅ Aucun doublon')}")

    total_doublons += len(vrais_doublons)

    # ── 3. Multi-parcelles légitimes ────────────────────────
    multi = {(nom, usine): entries for (nom, usine), entries in groups.items()
             if len(entries) > 1 and not any(
                 Counter(e["ton"] for e in entries)[t] >= 2
                 for t in Counter(e["ton"] for e in entries))}
    if multi:
        print(f"     Multi-parcelles légitimes ({len(multi)}):")
        for (nom, usine), entries in list(multi.items())[:5]:
            tons = sorted([e["ton"] for e in entries], reverse=True)
            print(f"        ✅ {nom} → {usine}: {[f'{t:,.0f}t' for t in tons]}")
        if len(multi) > 5:
            print(f"        ... et {len(multi)-5} autres")

    # ── 4. Tonnages aberrants ────────────────────────────────
    aberrants = []
    for r in rows:
        ton = float(r.get("tonnage_total", 0) or 0)
        nom = r.get("nom", "?")
        if ton <= 0:
            aberrants.append((nom, ton, "ZÉRO ou NÉGATIF"))
            problems.append(("TONNAGE_NUL", comm, nom, r.get("usine","?"), f"{ton}t"))
        elif ton < MIN_TONNAGE_PER_FARMER:
            aberrants.append((nom, ton, f"< {MIN_TONNAGE_PER_FARMER}t (suspect)"))
            problems.append(("TONNAGE_FAIBLE", comm, nom, r.get("usine","?"), f"{ton}t"))
        elif ton > MAX_TONNAGE_PER_FARMER:
            aberrants.append((nom, ton, f"> {MAX_TONNAGE_PER_FARMER:,}t (suspect)"))
            problems.append(("TONNAGE_ELEVE", comm, nom, r.get("usine","?"), f"{ton}t"))

    if aberrants:
        print(f"     {yellow(f'⚠️  {len(aberrants)} tonnage(s) suspect(s):')}")
        for nom, ton, raison in aberrants[:5]:
            print(f"        → {nom}: {ton:,.0f}t — {raison}")
    else:
        print(f"     Tonnages : {green('✅ Tous normaux')}")

    # ── 5. Données manquantes (région/zone vide) ─────────────
    sans_region = [r for r in rows if not r.get("region","")]
    sans_usine  = [r for r in rows if not r.get("usine","")]
    if sans_region:
        print(f"     {yellow(f'⚠️  {len(sans_region)} agriculteur(s) sans région:')}")
        for r in sans_region[:3]:
            print(f"        → {r.get('nom','?')} (zone: {r.get('zone','?')})")
        problems.append(("SANS_REGION", comm, f"{len(sans_region)} agriculteurs", "-", "région vide"))
    if sans_usine:
        print(f"     {red(f'❌ {len(sans_usine)} agriculteur(s) sans usine assignée')}")
        for r in sans_usine[:3]:
            print(f"        → {r.get('nom','?')}")
        problems.append(("SANS_USINE", comm, f"{len(sans_usine)} agriculteurs", "-", "usine vide"))

    # ── 6. Top 5 plus gros agriculteurs ─────────────────────
    sorted_rows = sorted(rows, key=lambda r: float(r.get("tonnage_total",0) or 0), reverse=True)
    print(f"     Top 5    :")
    for r in sorted_rows[:5]:
        ton = float(r.get("tonnage_total",0) or 0)
        pct = ton/actual*100 if actual > 0 else 0
        print(f"        {r.get('nom','?'):<28} {ton:>7,.0f}t ({pct:.1f}%) → {r.get('usine','?')}")

    all_problems.extend(problems)
    print()

# ── Résumé global ─────────────────────────────────────────────
print(bold("="*65))
print(bold("  RÉSUMÉ GLOBAL"))
print(bold("="*65))
print(f"  Total réel en base : {bold(f'{total_reel:,.0f}t')}")
print(f"  Total attendu      : {TOTAL_EXPECTED:,.0f}t")
ecart_global = total_reel - TOTAL_EXPECTED
if abs(ecart_global) < 50:
    print(f"  Écart global       : {green(f'{ecart_global:+,.0f}t ✅ PARFAIT')}")
elif ecart_global > 0:
    print(f"  Écart global       : {yellow(f'{ecart_global:+,.0f}t ⚠️ SURPLUS')}")
else:
    print(f"  Écart global       : {red(f'{ecart_global:+,.0f}t ❌ MANQUE')}")

print()
if not all_problems:
    print(f"  {green('✅ AUCUN PROBLÈME DÉTECTÉ — données parfaites !')}")
else:
    by_type = Counter(p[0] for p in all_problems)
    print(f"  Problèmes détectés ({len(all_problems)}) :")
    labels = {
        "DOUBLON":       "❌ Vrais doublons",
        "SURPLUS":       "⚠️  Surplus tonnage",
        "MANQUE":        "❌ Manque tonnage",
        "TONNAGE_NUL":   "❌ Tonnage nul",
        "TONNAGE_FAIBLE":"⚠️  Tonnage faible",
        "TONNAGE_ELEVE": "⚠️  Tonnage élevé",
        "SANS_REGION":   "⚠️  Sans région",
        "SANS_USINE":    "❌ Sans usine",
        "VIDE":          "❌ Commercial vide",
    }
    for ptype, cnt in sorted(by_type.items()):
        print(f"    {labels.get(ptype, ptype):<30} : {cnt}")

# Compter les doublons possibles depuis problems
n_doublons_possibles = sum(1 for p in all_problems if p[0] == "DOUBLON_POSSIBLE")
if n_doublons_possibles > 0:
    print(f"\n  {yellow('⚠️  Action manuelle requise pour les doublons possibles:')}")
    print(f"  1. Ouvre le dashboard → Gestion Agriculteurs")
    print(f"  2. Cherche l'agriculteur signalé")
    print(f"  3. Si upload en double → supprimer la ligne en double")
    print(f"  4. Si 2 vraies parcelles → ignorer l'alerte")
    print(f"  5. Relancer: python optimizer_v2.py && python migrate.py")

# ── Export CSV ────────────────────────────────────────────────
if args.export and all_problems:
    import csv
    fname = f"rapport_verification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(fname, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Type", "Commercial", "Agriculteur", "Usine", "Détail"])
        w.writerows(all_problems)
    print(f"\n  📄 Rapport exporté : {fname}")

print(bold("\n" + "="*65))