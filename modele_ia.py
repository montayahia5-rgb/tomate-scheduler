# -*- coding: utf-8 -*-
"""
modele_ia.py — Modèle IA de prédiction des tonnages par région × usine
=======================================================================
FONCTIONNEMENT:
  1. Lit les 3 prévisions (Déc25, Mai26, Juin26) + données historiques 2025
  2. Applique une moyenne pondérée temporelle (plus récent = plus fiable)
  3. Vérifie la compatibilité avec tous les caps (usines + commerciaux)
  4. Génère une prédiction finale optimisée

QUAND AJOUTER DE NOUVELLES DONNÉES:
  - Ajoute une entrée dans PREVISIONS avec la nouvelle date
  - Mets à jour WEIGHTS avec la nouvelle pondération
  - Lance: python modele_ia.py

USAGE:
  python modele_ia.py                    → prédiction + rapport
  python modele_ia.py --export excel     → génère Excel
  python modele_ia.py --export supabase  → pousse dans Supabase
"""

import sys, math, json, os
import pandas as pd
import numpy as np

# ============================================================
# CONFIGURATION — à mettre à jour à chaque nouvelle prévision
# ============================================================

REGIONS = ['CAP BON','NORD','GAFSA / KASSRINE','KAIROUAN','SIDI BOUZID','BOUFICHA']
USINES  = ['SICAM','TUCAL','COMOCAP','ABIDA','ELFALLEH']
SAISON_JOURS = 78   # durée saison en jours (15 juin → 31 août)

# Caps journaliers (tonnes/jour)
FACTORY_CAPS = {
    'SICAM':    1300,
    'TUCAL':     750,
    'COMOCAP':   700,
    'ABIDA':     150,
    'ELFALLEH':  100,
}
COMMERCIAL_CAPS = {
    'FEDI':             850,
    'MAKKI BEN SALAH':  800,
    'KHALIL':           800,
    'ACHREF AJLANI':    500,
    'JILANI OBAY':       50,   # cap officiel 50t/j
}

# ============================================================
# DONNÉES DES PRÉVISIONS — source vérifiée
# ============================================================
# Structure: tonnage = production prévue par région
#            SICAM/TUCAL/... = besoin de l'usine depuis cette région
#            BESOIN = total besoin usines = somme des colonnes usines

PREVISIONS = {
    'DEC25': {
        'date': '2026-01-16',
        'fiabilite': 0.20,   # 20% — prévision précoce, moins précise
        'CAP BON':          {'tonnage':48000,'SICAM':22000,'TUCAL':10000,'COMOCAP':13000,'ABIDA':0,   'ELFALLEH':3000,'BESOIN':48000},
        'NORD':             {'tonnage':12000,'SICAM':3500, 'TUCAL':7000, 'COMOCAP':1500, 'ABIDA':500, 'ELFALLEH':500, 'BESOIN':13000},
        'GAFSA / KASSRINE': {'tonnage':15500,'SICAM':11500,'TUCAL':2500, 'COMOCAP':1500, 'ABIDA':2500,'ELFALLEH':0,   'BESOIN':18000},
        'KAIROUAN':         {'tonnage':9000, 'SICAM':4500, 'TUCAL':2000, 'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':10000},
        'SIDI BOUZID':      {'tonnage':5500, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3500, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':500, 'BESOIN':3500},
    },
    'MAI26': {
        'date': '2026-05-09',
        'fiabilite': 0.35,   # 35% — prévision intermédiaire
        'CAP BON':          {'tonnage':49000,'SICAM':22000,'TUCAL':11000,'COMOCAP':13000,'ABIDA':0,   'ELFALLEH':4000,'BESOIN':50000},
        'NORD':             {'tonnage':12500,'SICAM':3500, 'TUCAL':4500, 'COMOCAP':3000, 'ABIDA':1000,'ELFALLEH':500, 'BESOIN':12500},
        'GAFSA / KASSRINE': {'tonnage':17000,'SICAM':11000,'TUCAL':2000, 'COMOCAP':500,  'ABIDA':3500,'ELFALLEH':0,   'BESOIN':17000},
        'KAIROUAN':         {'tonnage':8000, 'SICAM':4500, 'TUCAL':1200, 'COMOCAP':1000, 'ABIDA':1500,'ELFALLEH':0,   'BESOIN':8200},
        'SIDI BOUZID':      {'tonnage':4000, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3000, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':0,   'BESOIN':3000},
    },
    'JUN26': {
        'date': '2026-06-03',
        'fiabilite': 0.45,   # 45% — prévision la plus récente = plus fiable
        'CAP BON':          {'tonnage':49000,'SICAM':22000,'TUCAL':11000,'COMOCAP':12500,'ABIDA':0,   'ELFALLEH':4000,'BESOIN':49500},
        'NORD':             {'tonnage':12500,'SICAM':3500, 'TUCAL':4500, 'COMOCAP':3000, 'ABIDA':1000,'ELFALLEH':500, 'BESOIN':12500},
        'GAFSA / KASSRINE': {'tonnage':18000,'SICAM':11000,'TUCAL':2500, 'COMOCAP':500,  'ABIDA':3500,'ELFALLEH':0,   'BESOIN':17500},
        'KAIROUAN':         {'tonnage':9000, 'SICAM':4500, 'TUCAL':1200, 'COMOCAP':1500, 'ABIDA':1500,'ELFALLEH':500, 'BESOIN':9200},
        'SIDI BOUZID':      {'tonnage':6500, 'SICAM':2500, 'TUCAL':500,  'COMOCAP':1500, 'ABIDA':2000,'ELFALLEH':0,   'BESOIN':6500},
        'BOUFICHA':         {'tonnage':3000, 'SICAM':1000, 'TUCAL':1000, 'COMOCAP':1000, 'ABIDA':0,   'ELFALLEH':0,   'BESOIN':3000},
    },
    # ── AJOUTER LA PROCHAINE PRÉVISION ICI ──────────────────
    # 'JUL26': {
    #     'date': '2026-07-01',
    #     'fiabilite': 0.55,   # augmenter si plus récente
    #     'CAP BON': {...},
    #     ...
    # },
}

# ============================================================
# MODÈLE PRINCIPAL
# ============================================================

def predict_all():
    """
    Calcule la prédiction finale pour toutes les régions × usines.
    
    Algorithme:
    1. Moyenne pondérée par fiabilité temporelle
    2. Normalisation pour conserver le tonnage prédit
    3. Arrondi au 100 supérieur
    4. Vérification des contraintes
    5. Ajustement si dépassement de cap
    """
    prev_names = list(PREVISIONS.keys())
    predictions = {}

    for region in REGIONS:
        # Vérification que les fiabilités somment à 1
        total_w = sum(PREVISIONS[p]['fiabilite'] for p in prev_names)

        # Tonnage prédit = moyenne pondérée
        tonnage_pred = sum(
            PREVISIONS[p]['fiabilite'] / total_w * PREVISIONS[p][region]['tonnage']
            for p in prev_names
        )

        # Besoins usines prédits = moyenne pondérée par fiabilité
        usine_pred = {}
        for u in USINES:
            usine_pred[u] = sum(
                PREVISIONS[p]['fiabilite'] / total_w * PREVISIONS[p][region][u]
                for p in prev_names
            )

        # Normalisation: ajuster pour que total besoins = tonnage prédit
        total_b = sum(usine_pred[u] for u in USINES)
        if total_b > 0:
            scale = tonnage_pred / total_b
            usine_norm = {u: usine_pred[u] * scale for u in USINES}
        else:
            usine_norm = usine_pred.copy()

        # Arrondi au 100 supérieur
        usine_rounded = {u: math.ceil(usine_norm[u] / 100) * 100 for u in USINES}

        # Vérification caps journaliers
        warnings = []
        for u in USINES:
            daily = usine_rounded[u] / SAISON_JOURS
            if daily > FACTORY_CAPS[u]:
                warnings.append(
                    f"{u}: {daily:.0f}t/j dépasse le cap de {FACTORY_CAPS[u]}t/j"
                )
                # Ajustement: plafonner au cap × jours
                usine_rounded[u] = FACTORY_CAPS[u] * SAISON_JOURS

        total_final = sum(usine_rounded[u] for u in USINES)

        predictions[region] = {
            'tonnage_prevu':     round(tonnage_pred / 100) * 100,
            'tonnage_final':     total_final,
            'usines':            usine_rounded,
            'warnings':          warnings,
            'convergence':       _calc_convergence(region, prev_names),
        }

    return predictions


def _calc_convergence(region, prev_names):
    """
    Calcule l'indice de convergence des prévisions.
    Si Déc→Mai→Juin convergent (même direction) = haute confiance.
    Si elles divergent = basse confiance.
    """
    tonnages = [PREVISIONS[p][region]['tonnage'] for p in prev_names]
    if len(tonnages) < 2:
        return 'Indéterminé'
    deltas = [tonnages[i+1] - tonnages[i] for i in range(len(tonnages)-1)]
    if all(d >= 0 for d in deltas):
        return '🟢 Haute — tendance croissante stable'
    elif all(d <= 0 for d in deltas):
        return '🟢 Haute — tendance décroissante stable'
    else:
        return '🟡 Moyenne — prévisions en révision'


def verify_commercial_compatibility(predictions):
    """Vérifie que les prédictions sont compatibles avec les commerciaux."""
    COMM_DATA = {
        'FEDI':            {'tonnage': 33712, 'regions': ['CAP BON']},
        'MAKKI BEN SALAH': {'tonnage': 21265, 'regions': ['CAP BON','NORD','KAIROUAN']},
        'KHALIL':          {'tonnage': 15120, 'regions': ['NORD','KAIROUAN','SIDI BOUZID']},
        'ACHREF AJLANI':   {'tonnage': 16625, 'regions': ['GAFSA / KASSRINE']},
        'JILANI OBAY':     {'tonnage':  6965, 'regions': ['NORD']},
    }
    results = []
    total_actual = sum(c['tonnage'] for c in COMM_DATA.values())
    total_pred   = sum(predictions[r]['tonnage_final'] for r in REGIONS)

    for comm, data in COMM_DATA.items():
        daily = data['tonnage'] / SAISON_JOURS
        cap   = COMMERCIAL_CAPS[comm]
        share = data['tonnage'] / total_actual
        alloc = round(total_pred * share / 100) * 100
        results.append({
            'commercial':     comm,
            'tonnage_actuel': data['tonnage'],
            'tonnage_alloue': alloc,
            'cap_journalier': cap,
            'besoin_journalier': round(daily, 0),
            'ok': daily <= cap,
        })
    return results


def print_report(predictions, comm_results):
    """Affiche le rapport complet."""
    print()
    print("="*70)
    print("  MODÈLE IA — PRÉDICTIONS FINALES SAISON 2026")
    print("="*70)
    print(f"  Prévisions utilisées: {', '.join(PREVISIONS.keys())}")
    weights_str = ', '.join(f'{k}={v["fiabilite"]*100:.0f}%' for k,v in PREVISIONS.items())
    print(f"  Poids: {weights_str}")
    print()

    print(f"{'RÉGION':<22} {'TONNAGE':>10} {'SICAM':>8} {'COMOCAP':>8} {'TUCAL':>7} {'ABIDA':>7} {'ELFALLEH':>9} {'CONFIANCE'}")
    print("-"*95)
    for r in REGIONS:
        p = predictions[r]
        u = p['usines']
        print(f"{r:<22} {p['tonnage_final']:>10,} {u['SICAM']:>8,} {u['COMOCAP']:>8,} "
              f"{u['TUCAL']:>7,} {u['ABIDA']:>7,} {u['ELFALLEH']:>9,}  {p['convergence'][:20]}")
    print("-"*95)
    tt = sum(predictions[r]['tonnage_final'] for r in REGIONS)
    ts = sum(predictions[r]['usines']['SICAM'] for r in REGIONS)
    tc = sum(predictions[r]['usines']['COMOCAP'] for r in REGIONS)
    tu = sum(predictions[r]['usines']['TUCAL'] for r in REGIONS)
    ta = sum(predictions[r]['usines']['ABIDA'] for r in REGIONS)
    te = sum(predictions[r]['usines']['ELFALLEH'] for r in REGIONS)
    print(f"{'TOTAL':<22} {tt:>10,} {ts:>8,} {tc:>8,} {tu:>7,} {ta:>7,} {te:>9,}")

    print()
    print("  CAPS USINES (total / 78 jours):")
    all_ok = True
    for u in USINES:
        total_u = sum(predictions[r]['usines'][u] for r in REGIONS)
        daily   = total_u / SAISON_JOURS
        cap     = FACTORY_CAPS[u]
        ok      = daily <= cap
        status  = "✅" if ok else "❌"
        if not ok: all_ok = False
        print(f"    {u:<12}: {total_u:,}t | {daily:.0f}t/j | cap={cap}t/j {status}")
    if all_ok:
        print("    ✅ Tous les caps respectés")

    print()
    print("  COMPATIBILITÉ COMMERCIAUX:")
    for c in comm_results:
        status = "✅" if c['ok'] else "⚠️"
        print(f"    {c['commercial']:<20}: {c['tonnage_actuel']:,}t | "
              f"{c['besoin_journalier']:.0f}t/j vs cap {c['cap_journalier']}t/j {status}")

    print()
    print("  AVERTISSEMENTS:")
    any_warn = False
    for r in REGIONS:
        for w in predictions[r]['warnings']:
            print(f"    ⚠️  {r}: {w}")
            any_warn = True
    if not any_warn:
        print("    ✅ Aucun avertissement")
    print("="*70)


# ============================================================
# EXPORT
# ============================================================

def save_predictions(predictions):
    """Sauvegarde les prédictions en JSON pour réutilisation."""
    output = {
        'predictions': predictions,
        'factory_caps': FACTORY_CAPS,
        'commercial_caps': COMMERCIAL_CAPS,
        'previsions_used': list(PREVISIONS.keys()),
        'regions': REGIONS,
        'usines': USINES,
        'saison_jours': SAISON_JOURS,
    }
    with open('model_predictions.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("  Sauvegardé: model_predictions.json")


def push_to_supabase(predictions):
    """Pousse les prédictions dans Supabase (table region_predictions)."""
    try:
        from supabase import create_client
        url = os.environ.get('SUPABASE_URL', 'https://mwjefdqfzrtsfzspeppg.supabase.co')
        key = os.environ.get('SUPABASE_KEY', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44')
        sb = create_client(url, key)

        rows = []
        for region, p in predictions.items():
            for u in USINES:
                rows.append({
                    'region':  region,
                    'usine':   u,
                    'tonnage': int(p['usines'][u]),
                    'saison':  2026,
                    'source':  'modele_ia',
                })

        # Upsert
        sb.table('region_predictions').upsert(rows).execute()
        print(f"  ✅ {len(rows)} lignes poussées dans Supabase (table: region_predictions)")
    except Exception as e:
        print(f"  ❌ Erreur Supabase: {e}")


# ============================================================
# MAIN
# ============================================================

if __name__ == '__main__':
    print("Calcul des prédictions...")
    predictions  = predict_all()
    comm_results = verify_commercial_compatibility(predictions)

    print_report(predictions, comm_results)
    save_predictions(predictions)

    # Export selon argument
    if '--export' in sys.argv:
        mode = sys.argv[sys.argv.index('--export') + 1]
        if mode == 'excel':
            print("\nGénération Excel...")
            # Import et lance le générateur Excel
            os.system('python optimizer_v2.py')
            os.system('python migrate.py')
        elif mode == 'supabase':
            print("\nPush Supabase...")
            push_to_supabase(predictions)

    print("\nUtilisation:")
    print("  python modele_ia.py                    → rapport seul")
    print("  python modele_ia.py --export excel     → planning complet")
    print("  python modele_ia.py --export supabase  → table region_predictions")