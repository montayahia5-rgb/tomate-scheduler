# -*- coding: utf-8 -*-
"""
fix_all.py — Applique TOUS les fixes manquants sur optimizer_v2.py
Usage: python fix_all.py
"""
import re

with open('optimizer_v2.py', 'r', encoding='utf-8') as f:
    code = f.read()

original = code
fixes_applied = []

# ─── FIX 1: _is_rm=True dans rm_fixed_days ───────────────────────────────
if '"_is_rm":        True' not in code:
    # Chercher "Note": "RM-pre-alloc", suivi de }) et ajouter _is_rm
    code = re.sub(
        r'("Note":\s*"RM-pre-alloc",\s*\n(\s*))\}',
        r'\1    "_is_rm":        True,\n\2}',
        code, count=1
    )
    if '"_is_rm":        True' in code:
        fixes_applied.append('Fix1: _is_rm=True dans rm_fixed_days')
    else:
        print("  ❌ Fix1 ECHEC")
else:
    fixes_applied.append('Fix1: deja present')

# ─── FIX 2: _is_rm=False dans non-RM all_days ────────────────────────────
if '"_is_rm":        False' not in code:
    # Chercher "Note": note, suivi de }) puis FUSION RM
    code = re.sub(
        r'("Note":\s*note,\s*\n(\s*))\}(\s*\n\s*#[^\n]*FUSION)',
        r'\1    "_is_rm":        False,\n\2}\3',
        code, count=1
    )
    if '"_is_rm":        False' in code:
        fixes_applied.append('Fix2: _is_rm=False dans all_days non-RM')
    else:
        print("  ❌ Fix2 ECHEC")
else:
    fixes_applied.append('Fix2: deja present')

# ─── FIX 3: Skip RM dans boucle jours doubles ────────────────────────────
if 'if row.get("_is_rm", False):\n        continue' not in code:
    # Chercher nb_jours_doubles = 0 puis la boucle for row in all_days
    code = re.sub(
        r'(nb_jours_doubles = 0\s*\nfor row in all_days:\s*\n)(\s*key)',
        r'\1    if row.get("_is_rm", False):\n        continue\n\2key',
        code, count=1
    )
    if 'if row.get("_is_rm", False):\n        continue' in code:
        fixes_applied.append('Fix3: skip RM dans jours doubles')
    else:
        print("  ❌ Fix3 ECHEC")
else:
    fixes_applied.append('Fix3: deja present')

# ─── FIX 4: Filtre agri_indices par _is_rm ───────────────────────────────
if 'not r.get("_is_rm", False)' not in code:
    code = code.replace(
        'and r.get("Note") != "RM-pre-alloc"]',
        'and not r.get("_is_rm", False)]',
        1
    )
    if 'not r.get("_is_rm", False)' in code:
        fixes_applied.append('Fix4: filtre correction par _is_rm')
    else:
        print("  ❌ Fix4 ECHEC")
else:
    fixes_applied.append('Fix4: deja present')

# ─── FIX 5: Excel colonnes (eviter Length mismatch) ─────────────────────
if 'result_df[COLS1].copy()' not in code:
    code = code.replace(
        'disp = result_df.copy(); disp.columns = COLS1',
        'disp = result_df[COLS1].copy()',
        1
    )
    if 'result_df[COLS1].copy()' in code:
        fixes_applied.append('Fix5: Excel result_df[COLS1].copy()')
    else:
        print("  ❌ Fix5 ECHEC")
else:
    fixes_applied.append('Fix5: deja present')

# ─── Sauvegarde ──────────────────────────────────────────────────────────
if code != original:
    with open('optimizer_v2.py', 'w', encoding='utf-8') as f:
        f.write(code)
    print(f"✅ Fichier mis a jour ({len(fixes_applied)} fixes)")
else:
    print("ℹ️  Aucun changement (tous deja presents)")

for fix in fixes_applied:
    print(f"  ✅ {fix}")

# Verification finale
print("\nVerification:")
checks = [
    ('"_is_rm":        True',                   '_is_rm=True'),
    ('"_is_rm":        False',                  '_is_rm=False'),
    ('if row.get("_is_rm", False):\n        continue', 'skip RM jours doubles'),
    ('not r.get("_is_rm", False)',              'filtre _is_rm'),
    ('result_df[COLS1].copy()',                 'Excel COLS1'),
]
all_ok = True
for pattern, name in checks:
    ok = pattern in code
    print(f"  {'OK' if ok else 'MANQUE'} {name}")
    if not ok: all_ok = False
print("\n" + ("TOUS LES FIXES OK" if all_ok else "CERTAINS FIXES MANQUENT"))