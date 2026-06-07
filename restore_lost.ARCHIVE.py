# -*- coding: utf-8 -*-
"""
restore_lost.py — Restaure les 5 doublons KNOWN supprimés par --fix
====================================================================
Le --fix précédent a supprimé 10 doublons total, dont nous avons gardé
trace de 5 dans les logs (les plus gros) :

  KHALIL: AMOR KHECHIN → SICAM (630t)
  KHALIL: EZZEDINE GUESMI → SICAM (438t)
  KHALIL: EZZEDINE GUESMI → COMOCAP (262t)
  KHALIL: EZZEDINE GUESMI → ABIDA (350t)
  MAKKI BEN SALAH: KHALED CHATER → SICAM (350t)

Total restauré: 2,030t (KHALIL: 1,680t + MAKKI: 350t)

Reste perdus (à compléter manuellement) :
- ACHRAF: 2,572t (5-6 lignes inconnues)
- MAKKI:  490t additionnels

Utilisation: python restore_lost.py
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from supabase import create_client

SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

sb = create_client(SUPABASE_URL, SUPABASE_KEY)

# ── Données à restaurer (les 5 doublons connus) ────────────────────
LOST_ROWS = [
    {
        "commercial":    "KHALIL",
        "nom":           "AMOR KHECHIN",
        "region":        "KAIROUAN",
        "zone":          "BATTEN",
        "usine":         "SICAM",
        "accessibilite": "PL/PPL",
        "tonnage_total": 630.0,
        "date_debut":    "2026-07-01",
        "date_fin":      "2026-07-31",
    },
    {
        "commercial":    "KHALIL",
        "nom":           "EZZEDINE GUESMI",
        "region":        "KAIROUAN",
        "zone":          "SBIKHA-CHRARDA",
        "usine":         "SICAM",
        "accessibilite": "PL/PPL",
        "tonnage_total": 438.0,
        "date_debut":    "2026-07-01",
        "date_fin":      "2026-07-31",
    },
    {
        "commercial":    "KHALIL",
        "nom":           "EZZEDINE GUESMI",
        "region":        "KAIROUAN",
        "zone":          "SBIKHA-CHRARDA",
        "usine":         "COMOCAP",
        "accessibilite": "PL/PPL",
        "tonnage_total": 262.0,
        "date_debut":    "2026-07-01",
        "date_fin":      "2026-07-31",
    },
    {
        "commercial":    "KHALIL",
        "nom":           "EZZEDINE GUESMI",
        "region":        "KAIROUAN",
        "zone":          "SBIKHA-CHRARDA",
        "usine":         "ABIDA",
        "accessibilite": "PL/PPL",
        "tonnage_total": 350.0,
        "date_debut":    "2026-07-01",
        "date_fin":      "2026-07-31",
    },
    {
        "commercial":    "MAKKI BEN SALAH",
        "nom":           "KHALED CHATER",
        "region":        "CAP BON 2",
        "zone":          "MENZEL TAMIM",
        "usine":         "SICAM",
        "accessibilite": "PL/PPL",
        "tonnage_total": 350.0,
        "date_debut":    "2026-07-01",
        "date_fin":      "2026-07-31",
    },
]

print("="*60)
print("  RESTAURATION DES 5 DOUBLONS CONNUS SUPPRIMÉS PAR --FIX")
print("="*60)

# Vérifier qu'ils ne sont pas déjà présents (éviter de RE-créer doublons)
restored = 0
skipped = 0
for row in LOST_ROWS:
    # Chercher s'il existe DÉJÀ une ligne avec ce tonnage exact
    existing = sb.table("agriculteurs").select("id").eq(
        "commercial", row["commercial"]).eq(
        "nom", row["nom"]).eq(
        "usine", row["usine"]).eq(
        "tonnage_total", row["tonnage_total"]).execute().data
    
    if existing:
        print(f"  ⏭️  {row['commercial']} | {row['nom']} → {row['usine']} ({row['tonnage_total']}t) — DÉJÀ présent, skip")
        skipped += 1
        continue
    
    # Insérer
    sb.table("agriculteurs").insert(row).execute()
    print(f"  ✅ {row['commercial']} | {row['nom']} → {row['usine']} ({row['tonnage_total']}t) RESTAURÉ")
    restored += 1

print()
print(f"  Restaurés: {restored} | Skippés (déjà présents): {skipped}")
print()

# Vérification finale
print("="*60)
print("  VÉRIFICATION FINALE")
print("="*60)
all_data = sb.table("agriculteurs").select("commercial,tonnage_total").execute().data
from collections import defaultdict
by_comm = defaultdict(float)
for r in all_data:
    by_comm[r["commercial"]] += float(r["tonnage_total"] or 0)

expected = {
    "FEDI":31690,"MAKKI BEN SALAH":21365,"KHALIL":15705,
    "ACHREF AJLANI":14140,"JILANI OBAY":7410
}
total = 0
for comm in sorted(by_comm.keys()):
    exp = expected.get(comm, 0)
    ecart = by_comm[comm] - exp
    ok = "✅" if abs(ecart) < 100 else ("⚠️" if abs(ecart) < 1000 else "❌")
    print(f"  {comm:<22}: {by_comm[comm]:>9,.0f}t (attendu {exp:>6,}t, écart {ecart:+,.0f}t) {ok}")
    total += by_comm[comm]
print(f"  {'TOTAL':<22}: {total:>9,.0f}t")
print()
print("📝 Notes:")
print("  - Après ce script, écart devrait passer de -4,289t à ~-2,259t")
print("  - Le reste (ACHRAF 2,572t + MAKKI 490t) nécessite info manuelle")
print()
print("Prochaine étape: python optimizer_v2.py && python migrate.py")