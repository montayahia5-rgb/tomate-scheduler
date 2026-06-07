# -*- coding: utf-8 -*-
"""
supabase_check.py — Outil de diagnostic et fiabilité Supabase
==============================================================
Lance depuis le terminal:
    python supabase_check.py
    python supabase_check.py --fix    ← tente de corriger les problèmes
    python supabase_check.py --watch  ← surveille en continu (10s)
"""

import sys, time, datetime, json, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from supabase import create_client

def fetch_all(sb, table, columns="*", filters=None):
    """
    Récupère TOUTES les lignes d'une table Supabase avec pagination.
    Supabase a une limite PAR DÉFAUT de 1000 lignes par requête — sans
    pagination, on rate les lignes au-delà.
    """
    all_rows = []
    page_size = 1000
    offset = 0
    while True:
        q = sb.table(table).select(columns)
        if filters:
            for col, val in filters.items():
                q = q.eq(col, val)
        batch = q.range(offset, offset + page_size - 1).execute().data
        if not batch:
            break
        all_rows.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size
    return all_rows


SUPABASE_URL = "https://mwjefdqfzrtsfzspeppg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im13amVmZHFmenJ0c2Z6c3BlcHBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAzODg3MDgsImV4cCI6MjA5NTk2NDcwOH0.H5IX2uLneHdvwyLJrgN7OrHGrLSZSrQBAuJeuzCEz44"

# Valeurs de référence attendues après migration complète
EXPECTED = {
    "agriculteurs": {
        "min_rows": 200,      # on a 227 agriculteurs
        "max_rows": 300,
        "min_tonnage": 85000, # 90,310t attendu
        "max_tonnage": 95000,
        "commerciaux": ["ACHREF AJLANI","FEDI","JILANI OBAY","KHALIL","MAKKI BEN SALAH"],
    },
    "planning": {
        "min_rows": 2000,   # minimum de lignes attendu (chaque agriculteur × jour actif)
        "max_rows": 4000,
        # Note: tonnage planning = tonnes/JOUR (pas saison), pas de vérif tonnage ici
    },
    "transport": {
        "min_rows": 100,
        "max_rows": 500,
    },
}

# ── Couleurs terminal ──────────────────────────────────────────────────
G = "\033[92m"  # vert
R = "\033[91m"  # rouge
Y = "\033[93m"  # jaune
B = "\033[94m"  # bleu
W = "\033[97m"  # blanc
E = "\033[0m"   # reset

def ok(msg):  print(f"  {G}✅ {msg}{E}")
def err(msg): print(f"  {R}❌ {msg}{E}")
def warn(msg):print(f"  {Y}⚠️  {msg}{E}")
def info(msg):print(f"  {B}ℹ️  {msg}{E}")

# ══════════════════════════════════════════════════════════════════════
def check_connection(sb):
    """Test 1: Connexion Supabase."""
    print(f"\n{W}[1] TEST CONNEXION{E}")
    try:
        r = sb.table("agriculteurs").select("id", count="exact").limit(1).execute()
        ok(f"Connexion OK | {r.count} agriculteurs trouvés")
        return True
    except Exception as e:
        err(f"Connexion ÉCHOUÉE: {e}")
        return False

# ══════════════════════════════════════════════════════════════════════
def check_table_integrity(sb):
    """Test 2: Intégrité de chaque table."""
    print(f"\n{W}[2] INTÉGRITÉ DES TABLES{E}")
    results = {}
    
    for table, exp in EXPECTED.items():
        try:
            # Count rows
            r = sb.table(table).select("*", count="exact").limit(1).execute()
            count = r.count
            data  = fetch_all(sb, table)  # paginé
            
            # Tonnage check (tables avec tonnage)
            tonnage = 0
            if table == "agriculteurs" and data:
                tonnage = sum(float(row.get("tonnage_total",0) or 0) for row in data)
            elif table == "planning" and data:
                # tonnes_jour = livraison du JOUR (pas saison entière)
                # Le total planning = nombre de lignes × tonnes moyennes/jour
                # On vérifie le nombre de jours uniques et le total
                dates_uniq = len(set(row.get("date","") for row in data if row.get("date")))
                tonnage = sum(float(row.get("tonnes_jour",0) or 0) for row in data)
                # Pour planning: vérifier min_rows seulement (pas le tonnage)
                tonnage = 0  # skip tonnage check for planning
            
            # Vérifications
            problems = []
            if count < exp.get("min_rows",0):
                problems.append(f"trop peu de lignes: {count} < {exp['min_rows']} attendu")
            if count > exp.get("max_rows",99999):
                problems.append(f"trop de lignes: {count} > {exp['max_rows']}")
            if tonnage > 0 and tonnage < exp.get("min_tonnage",0):
                problems.append(f"tonnage trop faible: {tonnage:,.0f}t < {exp['min_tonnage']:,}t")
            if tonnage > 0 and tonnage > exp.get("max_tonnage",999999):
                problems.append(f"tonnage trop élevé: {tonnage:,.0f}t > {exp['max_tonnage']:,}t")
            
            # Commerciaux check pour agriculteurs
            if table == "agriculteurs" and data and "commerciaux" in exp:
                found_comms = set(row.get("commercial","") for row in data)
                missing = set(exp["commerciaux"]) - found_comms
                if missing:
                    problems.append(f"commerciaux manquants: {missing}")
            
            results[table] = {"rows": count, "tonnage": tonnage, "problems": problems}
            
            if problems:
                err(f"{table}: {count} lignes | {tonnage:,.0f}t")
                for p in problems:
                    warn(f"  → {p}")
            else:
                extra = f" | {tonnage:,.0f}t" if tonnage > 0 else ""
                ok(f"{table}: {count} lignes{extra}")
                
        except Exception as e:
            err(f"{table}: ERREUR LECTURE → {e}")
            results[table] = {"error": str(e)}
    
    return results

# ══════════════════════════════════════════════════════════════════════
def check_data_freshness(sb):
    """Test 3: Fraîcheur des données (évite données périmées)."""
    print(f"\n{W}[3] FRAÎCHEUR DES DONNÉES{E}")
    
    try:
        # Vérifier la date max dans planning (avec pagination)
        plan = fetch_all(sb, "planning", "date")
        if not plan:
            warn("Planning vide — aucune date à vérifier")
            return
        
        dates = [row["date"] for row in plan if row.get("date")]
        if not dates:
            warn("Aucune date dans le planning")
            return
        
        dates.sort()
        date_min = dates[0]
        date_max = dates[-1]
        
        if date_max >= "2026-08-01":
            ok(f"Planning couvre jusqu'au {date_max} ✅")
        elif date_max >= "2026-07-15":
            warn(f"Planning s'arrête au {date_max} — devrait aller jusqu'en août")
            info("  → Cause probable: certains agriculteurs ont des fenêtres courtes")
            info("  → Solution: relancer python optimizer_v2.py puis python migrate.py")
        else:
            err(f"Planning trop court: jusqu'au {date_max} seulement")
            info("  → Solution: relancer python optimizer_v2.py puis python migrate.py")
        
        n_jours = len(set(dates))
        info(f"Période: {date_min} → {date_max} ({n_jours} jours uniques sur 67 attendus)")
        if n_jours < 60:
            warn(f"Seulement {n_jours}/67 jours couverts — planning incomplet")
        
        # Vérifier transport
        trans = fetch_all(sb, "transport", "date")
        if trans:
            t_dates = sorted(set(row["date"] for row in trans if row.get("date")))
            ok(f"Transport: {len(t_dates)} jours planifiés ({t_dates[0]} → {t_dates[-1]})")
        
    except Exception as e:
        err(f"Erreur vérification fraîcheur: {e}")

# ══════════════════════════════════════════════════════════════════════
def check_data_quality(sb):
    """Test 4: Qualité des données (doublons, nulls, régions invalides)."""
    print(f"\n{W}[4] QUALITÉ DES DONNÉES{E}")
    
    try:
        data = fetch_all(sb, "agriculteurs",
            "commercial,nom,tonnage_total,region,accessibilite")
        
        if not data:
            warn("Table agriculteurs vide")
            return
        
        total = len(data)
        
        # Lignes TOTAL (erreur connue)
        total_rows = [r for r in data if str(r.get("nom","")).upper().startswith("TOTAL")]
        if total_rows:
            err(f"{len(total_rows)} lignes TOTAL trouvées (à supprimer!)")
            for r in total_rows[:3]:
                print(f"     → nom='{r['nom']}' tonnage={r.get('tonnage_total')}")
        else:
            ok("Aucune ligne TOTAL parasite")
        
        # Régions invalides
        VALID_REGIONS = {"CAP BON 1","CAP BON 2","NORD","GAFSA / KASSRINE",
                         "KAIROUAN","SIDI BOUZID","BOUFICHA"}
        bad_reg = [r for r in data if r.get("region","") not in VALID_REGIONS]
        if bad_reg:
            bad_vals = list(set(r.get("region","?") for r in bad_reg))[:8]
            warn(f"{len(bad_reg)} lignes avec région invalide: {bad_vals}")
        else:
            ok("Toutes les régions sont valides")
        
        # Accessibilités invalides
        VALID_ACCESS = {"PL/PPL","PL/SEMI","RM","TRC/PPL","TRC/PPL/PL","PL/PPL/SEMI","PL"}
        bad_acc = [r for r in data if r.get("accessibilite","") not in VALID_ACCESS]
        if bad_acc:
            bad_vals = list(set(r.get("accessibilite","?") for r in bad_acc))[:5]
            warn(f"{len(bad_acc)} lignes avec accessibilité invalide: {bad_vals}")
        else:
            ok("Toutes les accessibilités sont valides")
        
        # Tonnages nuls
        null_ton = [r for r in data if not r.get("tonnage_total") or float(r.get("tonnage_total",0)) <= 0]
        if null_ton:
            warn(f"{len(null_ton)} lignes avec tonnage nul ou absent")
        else:
            ok("Tous les tonnages sont positifs")
        
        # Doublons (nom + usine) dans le même commercial
        # Note: (nom seul) peut apparaître plusieurs fois = normal si multi-usines
        # VRAI doublon = même nom + même usine + même commercial + même tonnage EXACT
        # Tonnages différents = multi-parcelles légitimes (ex: EZZEDINE GUESMI 1050t + 438t)
        from collections import Counter, defaultdict
        data_full = fetch_all(sb, "agriculteurs",
            "commercial,nom,usine,tonnage_total") or []
        
        # Grouper par (commercial, nom, usine) et lister les tonnages
        groups = defaultdict(list)
        for r in data_full:
            key = (r.get("commercial",""), r.get("nom",""), r.get("usine","") or "")
            groups[key].append(float(r.get("tonnage_total", 0) or 0))
        
        # Vrai doublon = même triplet ET même tonnage exact en double
        real_dups = {}
        for key, tonnages in groups.items():
            if key[2] == "":  # usine vide = pas un doublon
                continue
            ton_cnt = Counter(tonnages)
            dups_in_group = {t: n for t, n in ton_cnt.items() if n >= 2}
            if dups_in_group:
                real_dups[key] = (tonnages, dups_in_group)
        
        # Multi-parcelles légitimes (même triplet mais tonnages DIFFÉRENTS)
        multi_parcelles = {k: v for k, v in groups.items()
                          if len(v) > 1 and k[2] != "" and k not in real_dups}
        
        if real_dups:
            warn(f"{len(real_dups)} vrais doublons (même tonnage exact en double):")
            for (comm, nom, usine), (tonnages, dups_in) in list(real_dups.items())[:5]:
                print(f"     → {comm}: '{nom}' → {usine} × tonnages {[round(t,0) for t in tonnages]}")
            print(f"  Cause probable: fichier uploadé 2 fois")
            print(f"  Solution: SQL dans Supabase →")
            print(f"    DELETE FROM agriculteurs WHERE id NOT IN (")
            print(f"      SELECT MIN(id) FROM agriculteurs GROUP BY commercial, nom, usine, tonnage_total);")
        else:
            ok("Aucun doublon (nom+usine+commercial+tonnage) détecté")
        
        if multi_parcelles:
            info(f"{len(multi_parcelles)} agriculteurs multi-parcelles (tonnages différents = NORMAL):")
            for (comm, nom, usine), tonnages in list(multi_parcelles.items())[:3]:
                print(f"     ✅ {comm}: '{nom}' → {usine} × {[round(t,0) for t in tonnages]}")
        
        # Cas usine vide = données incomplètes (pas des doublons)
        empty_usine = [r for r in data_full if not r.get("usine")]
        if empty_usine:
            warn(f"{len(empty_usine)} agriculteurs sans usine renseignée dans Supabase")
        
        # Info: agriculteurs multi-usines (normal)
        nom_comm_only = [(r.get("commercial",""), r.get("nom","")) for r in data]
        multi_usine = {k:v for k,v in Counter(nom_comm_only).items() if v > 1}
        if multi_usine:
            info(f"{len(multi_usine)} agriculteurs livrent à plusieurs usines (normal)")
        
        info(f"Résumé: {total} lignes | {sum(float(r.get('tonnage_total',0)) for r in data):,.0f}t total")
        
    except Exception as e:
        err(f"Erreur vérification qualité: {e}")

# ══════════════════════════════════════════════════════════════════════
def test_write_read(sb):
    """Test 5: Écriture + relecture (fiabilité Supabase)."""
    print(f"\n{W}[5] TEST ÉCRITURE/LECTURE (fiabilité){E}")
    TEST_TABLE = "depot_status"
    TEST_COMM  = "_TEST_DIAGNOSTIC_"
    
    try:
        # Écrire
        ts = datetime.datetime.now().isoformat()
        sb.table(TEST_TABLE).upsert({
            "commercial":  TEST_COMM,
            "statut":      "test",
            "nb_agri":     42,
            "tonnage":     12345.6,
            "depose_le":   ts,
            "fichier_nom": "test_diagnostic.py",
        }).execute()
        
        # Relire immédiatement
        r = sb.table(TEST_TABLE).select("*").eq("commercial", TEST_COMM).execute().data
        
        if not r:
            err("ÉCRITURE OK mais LECTURE retourne vide → problème Supabase!")
        elif r[0]["nb_agri"] != 42:
            err(f"Données corrompues: nb_agri={r[0]['nb_agri']} (attendu 42)")
        else:
            ok(f"Écriture + relecture cohérente ✅ (timestamp={ts[:19]})")
        
        # Nettoyage
        sb.table(TEST_TABLE).delete().eq("commercial", TEST_COMM).execute()
        ok("Nettoyage test effectué")
        
    except Exception as e:
        err(f"Test écriture/lecture ÉCHOUÉ: {e}")

# ══════════════════════════════════════════════════════════════════════
def fix_known_issues(sb):
    """Mode --fix: corrige les problèmes connus automatiquement."""
    print(f"\n{W}[FIX] CORRECTION AUTOMATIQUE DES PROBLÈMES CONNUS{E}")
    
    try:
        # 0. Supprimer vrais doublons (garder le MIN(id) pour chaque triplet unique)
        # Note: Supabase ne supporte pas DELETE avec GROUP BY directement
        # On récupère les IDs à garder puis on supprime les autres
        all_agri = sb.table("agriculteurs").select("id,commercial,nom,usine").execute().data
        if all_agri:
            seen = {}
            ids_to_delete = []
            for r in sorted(all_agri, key=lambda x: x.get("id",0)):
                key = (r.get("commercial",""), r.get("nom",""), r.get("usine","") or "")
                if key in seen:
                    ids_to_delete.append(r["id"])
                else:
                    seen[key] = r["id"]
            if ids_to_delete:
                for id_del in ids_to_delete:
                    sb.table("agriculteurs").delete().eq("id", id_del).execute()
                ok(f"{len(ids_to_delete)} doublons supprimés")
            else:
                ok("Aucun doublon à supprimer")
        
        # 1. Supprimer lignes TOTAL
        r = sb.table("agriculteurs").delete().ilike("nom","TOTAL%").execute()
        ok(f"Lignes TOTAL supprimées")
        
        r = sb.table("agriculteurs").delete().ilike("nom","SOUS-TOTAL%").execute()
        ok(f"Lignes SOUS-TOTAL supprimées")
        
        # 2. Normaliser régions
        region_fixes = [
            ("region","nabeul","CAP BON 2"), ("region","NABEUL","CAP BON 2"),
            ("region","beja","NORD"),         ("region","BEJA","NORD"),
            ("region","manouba","NORD"),       ("region","MANOUBA","NORD"),
            ("region","gafsa","GAFSA / KASSRINE"),
            ("region","kassrine","GAFSA / KASSRINE"),
        ]
        for col, old_val, new_val in region_fixes:
            sb.table("agriculteurs").update({col: new_val}).ilike(col, old_val).execute()
        ok("Régions normalisées (nabeul→CAP BON 2, beja→NORD, etc.)")
        
        # 3. Supprimer tonnages nuls
        sb.table("agriculteurs").delete().lte("tonnage_total", 0).execute()
        ok("Lignes avec tonnage ≤ 0 supprimées")
        
    except Exception as e:
        err(f"Erreur correction: {e}")

# ══════════════════════════════════════════════════════════════════════
def print_summary(results):
    """Résumé final par commercial."""
    print(f"\n{W}[6] RÉSUMÉ PAR COMMERCIAL{E}")
    try:
        sb = create_client(SUPABASE_URL, SUPABASE_KEY)
        data = fetch_all(sb, "agriculteurs", "commercial,nom,tonnage_total")
        if not data: return
        
        from collections import defaultdict
        by_comm = defaultdict(lambda: {"rows":0,"tonnage":0.0,"farmers":set()})
        for r in data:
            if str(r.get("nom","")).upper().startswith("TOTAL"): continue
            c = r.get("commercial","?")
            by_comm[c]["rows"] += 1
            by_comm[c]["tonnage"] += float(r.get("tonnage_total",0) or 0)
            by_comm[c]["farmers"].add(r.get("nom",""))
        
        expected_tons = {
            "FEDI":31690,"MAKKI BEN SALAH":21365,"KHALIL":15705,
            "ACHREF AJLANI":14140,"JILANI OBAY":7410
        }
        
        print(f"  {'Commercial':<22} {'Lignes':>7} {'Agriculteurs':>13} {'Tonnage':>10} {'Attendu':>10} {'Écart':>8}")
        print(f"  {'─'*75}")
        total_tons = 0
        for comm in sorted(by_comm.keys()):
            d = by_comm[comm]
            exp = expected_tons.get(comm, "?")
            ecart = d["tonnage"] - exp if isinstance(exp,int) else 0
            status = G+"✅"+E if abs(ecart)<100 else (Y+"⚠️"+E if abs(ecart)<1000 else R+"❌"+E)
            print(f"  {comm:<22} {d['rows']:>7} {len(d['farmers']):>13} "
                  f"{d['tonnage']:>9,.0f}t {str(exp)+' t':>10} {ecart:>+8,.0f}t {status}")
            total_tons += d["tonnage"]
        print(f"  {'─'*75}")
        print(f"  {'TOTAL DÉCLARÉ':<22} {'':>7} {'':>13} {total_tons:>9,.0f}t")
        
        # Comparer avec le planifié (depuis table planning)
        try:
            plan_data = fetch_all(sb, "planning", "tonnes_jour,date")
            if plan_data:
                total_planifie = sum(float(r.get("tonnes_jour",0) or 0) for r in plan_data)
                ecart = total_planifie - total_tons
                ecart_pct = (ecart / total_tons * 100) if total_tons > 0 else 0
                print(f"  {'TOTAL PLANIFIÉ':<22} {'':>7} {'':>13} {total_planifie:>9,.0f}t")
                print(f"  {'ÉCART':<22} {'':>7} {'':>13} {ecart:>+9,.0f}t ({ecart_pct:+.1f}%)")
                print()
                if abs(ecart_pct) <= 5:
                    print(f"  {G}✅ Écart normal (≤5%) — dû à la tolérance OR-Tools{E}")
                elif abs(ecart_pct) <= 15:
                    print(f"  {Y}⚠️  Écart modéré ({ecart_pct:+.1f}%) — vérifier les fenêtres{E}")
                else:
                    # Écart très grand → planning incomplet dans Supabase
                    print(f"  {R}❌ Écart anormal ({ecart_pct:+.1f}%) → PLANNING INCOMPLET dans Supabase{E}")
                    # Calculer le nombre de jours dans le planning
                    dates_plan = set(r.get("date","") for r in plan_data if r.get("date"))
                    print(f"  {R}   Supabase a seulement {len(dates_plan)} jours planifiés{E}")
                    print(f"  {R}   Le planning complet devrait avoir ~67 jours{E}")
                    print()
                    print(f"  {W}  ► SOLUTION :{E}")
                    print(f"     1. Lance dans le terminal : python optimizer_v2.py")
                    print(f"     2. Ensuite              : python migrate.py")
                    print(f"     3. Relance ce check pour vérifier")
                print()
                print(f"  {B}ℹ️  EXPLICATION DE L'ÉCART :{E}")
                print(f"  Déclaré  = tonnage TOTAL que chaque agriculteur prévoit de produire")
                print(f"             sur toute sa saison (ce que les commerciaux ont uploadé)")
                print(f"  Planifié = tonnage qu'OR-Tools a RÉUSSI à placer dans le calendrier")
                print(f"             en respectant toutes les contraintes (caps, fenêtres...)")
                print(f"  Écart    = tonnes que le solver n'a pas pu placer (tolérance ±5%)")
                print(f"             ou tonnage hors fenêtre de maturité disponible")
        except Exception as e:
            warn(f"Impossible de lire le planning pour comparaison: {e}")
        
    except Exception as e:
        err(f"Erreur résumé: {e}")

# ══════════════════════════════════════════════════════════════════════
def run_all_checks():
    mode_fix   = "--fix"   in sys.argv
    mode_watch = "--watch" in sys.argv
    
    while True:
        print(f"\n{'═'*60}")
        print(f"  DIAGNOSTIC SUPABASE — {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"{'═'*60}")
        
        try:
            sb = create_client(SUPABASE_URL, SUPABASE_KEY)
        except Exception as e:
            err(f"Impossible de créer le client Supabase: {e}")
            break
        
        if not check_connection(sb):
            print(R+"  Connexion impossible — vérifiez votre réseau et la clé API"+E)
            if mode_watch:
                print(f"\n  Prochain check dans 10s...")
                time.sleep(10); continue
            break
        
        check_table_integrity(sb)
        check_data_freshness(sb)
        check_data_quality(sb)
        test_write_read(sb)
        
        if mode_fix:
            fix_known_issues(sb)
        
        print_summary(sb)
        
        if not mode_watch:
            break
        
        print(f"\n  Prochain check dans 10s... (Ctrl+C pour arrêter)")
        time.sleep(10)

if __name__ == "__main__":
    run_all_checks()