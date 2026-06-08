# -*- coding: utf-8 -*-
"""
centre_tab.py — Module session CENTRE (BACCARA, KERKOUANE, 428...)
================================================================
Permet à un centre de:
  - Voir/ajouter/modifier ses sous-agriculteurs
  - Upload un fichier Excel de ses sous-agriculteurs
  - Voir son planning interne (redistribution journalière)
  - Validation: somme sous-agriculteurs = tonnage centre dans agriculteurs

Utilisation depuis dashboard_phase10.py:
    if CURRENT_ROLE == "centre":
        from centre_tab import render_centre_dashboard
        render_centre_dashboard(sb, CURRENT_FILTER, CURRENT_NAME)
"""
import streamlit as st
import pandas as pd
import math
from datetime import datetime, date

def render_centre_dashboard(sb, centre_nom, display_name):
    """
    Affiche le dashboard d'un centre.
    
    Args:
        sb: client Supabase
        centre_nom: nom du centre (ex: "STE BACCARA")
        display_name: nom affiché (peut être identique)
    """
    st.title(f"🏭 Centre : {display_name}")
    
    # ── Récupérer le tonnage TOTAL alloué au centre depuis agriculteurs ─
    try:
        centre_master = sb.table("agriculteurs").select(
            "commercial,tonnage_total,usine,region,date_debut,date_fin,accessibilite"
        ).eq("nom", centre_nom).execute().data or []
    except Exception as e:
        st.error(f"Erreur connexion Supabase: {e}")
        return
    
    if not centre_master:
        st.error(f"❌ Centre '{centre_nom}' introuvable dans la table agriculteurs.")
        st.info("Le centre doit être déclaré par son commercial parent (ex: FEDI) avant utilisation.")
        return
    
    # Tonnage total alloué au centre (somme sur toutes ses entrées dans agriculteurs)
    total_alloue = sum(float(r.get("tonnage_total", 0) or 0) for r in centre_master)
    commercial_parent = centre_master[0].get("commercial", "?")
    
    # ── KPIs en haut ──────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Commercial parent", commercial_parent)
    col2.metric("Tonnage alloué", f"{total_alloue:,.0f}t")
    
    # Lire les sous-agriculteurs existants
    try:
        sub_agri = sb.table("centre_agriculteurs").select("*").eq(
            "centre_nom", centre_nom).execute().data or []
    except Exception as e:
        st.error(f"Table centre_agriculteurs introuvable. Exécute centre_setup.sql dans Supabase.")
        st.code("CREATE TABLE centre_agriculteurs (...)  -- voir centre_setup.sql")
        return
    
    df_sub = pd.DataFrame(sub_agri) if sub_agri else pd.DataFrame()
    total_sub = df_sub["tonnage_total"].sum() if not df_sub.empty else 0
    nb_sub = len(df_sub)
    
    col3.metric("Sous-agriculteurs", nb_sub)
    
    diff = total_alloue - total_sub
    if abs(diff) < 1:
        col4.metric("Reste à attribuer", f"{diff:,.0f}t", delta="✅ Équilibré")
    elif diff > 0:
        col4.metric("Reste à attribuer", f"{diff:,.0f}t", 
                    delta=f"⚠️ Manque", delta_color="inverse")
    else:
        col4.metric("Reste à attribuer", f"{diff:,.0f}t",
                    delta=f"❌ Dépassement", delta_color="inverse")
    
    st.markdown("---")
    
    # ── 3 Tabs : Liste, Upload, Planning interne ─────────────────
    tab1, tab2, tab3 = st.tabs([
        "📋 Mes sous-agriculteurs",
        "📤 Upload fichier Excel",
        "📅 Mon planning",
    ])
    
    # ═══════════════════════════════════════════════════════════════
    # TAB 1 : Liste et gestion des sous-agriculteurs
    # ═══════════════════════════════════════════════════════════════
    with tab1:
        st.subheader(f"📋 Sous-agriculteurs de {centre_nom}")
        
        if df_sub.empty:
            st.info("Aucun sous-agriculteur encore. Utilise 'Upload Excel' ou ajoute manuellement ci-dessous.")
        else:
            # Affichage tableau
            display_cols = ["agriculteur_nom", "tonnage_total", "usine", "region", 
                           "zone", "accessibilite", "date_debut", "date_fin"]
            display_cols = [c for c in display_cols if c in df_sub.columns]
            df_display = df_sub[display_cols].copy()
            df_display.columns = ["Agriculteur", "Tonnage", "Usine", "Région",
                                  "Zone", "Accès", "Date début", "Date fin"][:len(display_cols)]
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            # Bouton export CSV
            st.download_button(
                "⬇️ Exporter CSV",
                data=df_display.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"sous_agriculteurs_{centre_nom.replace(' ','_')}.csv",
                mime="text/csv",
            )
        
        st.markdown("---")
        st.subheader("➕ Ajouter un sous-agriculteur")
        
        with st.form("add_sub_agri", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom *", placeholder="Ahmed Ben Slim")
            tonnage = c2.number_input("Tonnage (t) *", min_value=0.0, step=10.0)
            usine = c3.selectbox("Usine", ["SICAM", "COMOCAP", "TUCAL", "ABIDA", "ELFALLEH"])
            
            c4, c5, c6 = st.columns(3)
            region = c4.selectbox("Région", ["CAP BON 1","CAP BON 2","NORD",
                                              "GAFSA / KASSRINE","KAIROUAN","SIDI BOUZID","BOUFICHA"])
            zone = c5.text_input("Zone", placeholder="DAR ALOUCH")
            access = c6.selectbox("Accessibilité", ["PL/PPL","PL/SEMI","RM","TRC/PPL","TRC/PPL/PL","PPL","PL","SEMI"])
            
            c7, c8 = st.columns(2)
            date_debut = c7.date_input("Date début", value=date(2026, 6, 20))
            date_fin = c8.date_input("Date fin", value=date(2026, 8, 25))
            
            submitted = st.form_submit_button("➕ Ajouter")
            if submitted:
                if not nom or tonnage <= 0:
                    st.error("Nom et tonnage obligatoires.")
                else:
                    # Vérifier qu'on ne dépasse pas le tonnage alloué
                    new_total = total_sub + tonnage
                    if new_total > total_alloue * 1.001:  # tolérance arrondi
                        st.error(f"❌ Dépassement: ajout {tonnage:,.0f}t → total {new_total:,.0f}t > "
                                f"alloué {total_alloue:,.0f}t")
                    else:
                        try:
                            sb.table("centre_agriculteurs").insert({
                                "centre_nom":        centre_nom,
                                "commercial_parent": commercial_parent,
                                "agriculteur_nom":   nom.strip(),
                                "tonnage_total":     float(tonnage),
                                "usine":             usine,
                                "region":            region,
                                "zone":              zone.strip() if zone else "",
                                "accessibilite":     access,
                                "date_debut":        date_debut.isoformat(),
                                "date_fin":          date_fin.isoformat(),
                            }).execute()
                            st.success(f"✅ {nom} ajouté ({tonnage:,.0f}t)")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur: {e}")
        
        # Suppression
        if not df_sub.empty:
            st.markdown("---")
            st.subheader("🗑️ Supprimer un sous-agriculteur")
            options = {f"{r['agriculteur_nom']} ({r['tonnage_total']:,.0f}t → {r['usine']})": r["id"]
                       for _, r in df_sub.iterrows()}
            to_delete = st.selectbox("Choisir", list(options.keys()))
            if st.button("🗑️ Supprimer"):
                try:
                    sb.table("centre_agriculteurs").delete().eq("id", options[to_delete]).execute()
                    st.success("Supprimé.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur: {e}")
    
    # ═══════════════════════════════════════════════════════════════
    # TAB 2 : Upload Excel
    # ═══════════════════════════════════════════════════════════════
    with tab2:
        st.subheader(f"📤 Upload fichier sous-agriculteurs — {centre_nom}")
        st.caption("Le fichier Excel doit contenir les colonnes: "
                   "AGRICULTEUR, TONNAGE, USINE, REGION, ZONE, ACCESSIBILITE, DATE_DEBUT, DATE_FIN")
        
        # Template à télécharger
        template_df = pd.DataFrame({
            "AGRICULTEUR":   ["Ahmed Ben Slim", "Karim Hamdi"],
            "TONNAGE":       [800, 600],
            "USINE":         ["SICAM", "COMOCAP"],
            "REGION":        ["CAP BON 1", "CAP BON 2"],
            "ZONE":          ["DAR ALOUCH", "MENZEL TAMIM"],
            "ACCESSIBILITE": ["PL/PPL", "PL/SEMI"],
            "DATE_DEBUT":    ["2026-06-20", "2026-06-25"],
            "DATE_FIN":      ["2026-08-25", "2026-08-25"],
        })
        
        import io
        buffer = io.BytesIO()
        template_df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)
        st.download_button(
            "⬇️ Télécharger template Excel",
            data=buffer,
            file_name=f"template_{centre_nom.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
        st.markdown("---")
        uploaded = st.file_uploader("Choisir un fichier Excel", type=["xlsx", "xls"], key=f"upload_{centre_nom}")
        
        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                df_upload.columns = [str(c).strip().upper() for c in df_upload.columns]
                
                st.write("**Aperçu du fichier:**")
                st.dataframe(df_upload.head(10), use_container_width=True)
                
                # Validation du tonnage total
                if "TONNAGE" not in df_upload.columns:
                    st.error("❌ Colonne 'TONNAGE' manquante.")
                else:
                    total_upload = df_upload["TONNAGE"].sum()
                    nb_lignes = len(df_upload)
                    
                    st.info(f"📊 {nb_lignes} agriculteurs | Total: {total_upload:,.0f}t | "
                            f"Alloué centre: {total_alloue:,.0f}t")
                    
                    diff = total_upload - total_alloue
                    if abs(diff) < 1:
                        st.success("✅ Total parfait — correspond au tonnage alloué")
                    elif diff > 0:
                        st.warning(f"⚠️ Dépassement {diff:+,.0f}t")
                    else:
                        st.warning(f"⚠️ Sous-total {diff:+,.0f}t (manque)")
                    
                    if st.button("✅ Importer ces données", type="primary"):
                        try:
                            # Supprimer anciens sous-agriculteurs
                            sb.table("centre_agriculteurs").delete().eq(
                                "centre_nom", centre_nom).execute()
                            
                            # Insérer les nouveaux
                            rows = []
                            for _, r in df_upload.iterrows():
                                rows.append({
                                    "centre_nom":        centre_nom,
                                    "commercial_parent": commercial_parent,
                                    "agriculteur_nom":   str(r.get("AGRICULTEUR", "")).strip(),
                                    "tonnage_total":     float(r.get("TONNAGE", 0) or 0),
                                    "usine":             str(r.get("USINE", "")).strip().upper(),
                                    "region":            str(r.get("REGION", "")).strip().upper(),
                                    "zone":              str(r.get("ZONE", "")).strip(),
                                    "accessibilite":     str(r.get("ACCESSIBILITE", "PL/PPL")).strip(),
                                    "date_debut":        str(r.get("DATE_DEBUT", "2026-06-20"))[:10],
                                    "date_fin":          str(r.get("DATE_FIN", "2026-08-25"))[:10],
                                })
                            for i in range(0, len(rows), 100):
                                sb.table("centre_agriculteurs").insert(rows[i:i+100]).execute()
                            st.success(f"✅ {len(rows)} sous-agriculteurs importés!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur import: {e}")
            except Exception as e:
                st.error(f"Erreur lecture fichier: {e}")
    
    # ═══════════════════════════════════════════════════════════════
    # TAB 3 : Mon Planning (redistribution depuis planning global)
    # ═══════════════════════════════════════════════════════════════
    with tab3:
        st.subheader(f"📅 Mon planning — {centre_nom}")
        st.caption(f"Distribution journalière du tonnage alloué entre les sous-agriculteurs.")
        
        if df_sub.empty:
            st.info("Ajoute d'abord des sous-agriculteurs (Tab 'Mes sous-agriculteurs' ou 'Upload').")
        else:
            # Lire le planning global pour ce centre
            try:
                plan_master = sb.table("planning").select(
                    "date,tonnes_jour,usine,commercial"
                ).eq("agriculteur", centre_nom).execute().data or []
                
                if not plan_master:
                    st.warning(f"Aucun planning global trouvé pour {centre_nom}. "
                              f"Demande à l'admin de relancer optimizer_v2.py.")
                else:
                    df_plan = pd.DataFrame(plan_master)
                    df_plan["date"] = pd.to_datetime(df_plan["date"])
                    df_plan["tonnes_jour"] = pd.to_numeric(df_plan["tonnes_jour"], errors="coerce").fillna(0)
                    
                    # Stats du planning global
                    total_plan = df_plan["tonnes_jour"].sum()
                    nb_jours = df_plan["date"].nunique()
                    st.info(f"📊 Planning global: {nb_jours} jours | Total planifié: {total_plan:,.0f}t")
                    
                    # Distribution par jour : proratiser entre sous-agriculteurs
                    # Selon leur tonnage déclaré (poids)
                    sub_weights = df_sub.set_index("agriculteur_nom")["tonnage_total"].to_dict()
                    total_weight = sum(sub_weights.values())
                    
                    if total_weight == 0:
                        st.error("Total sous-agriculteurs = 0. Impossible de distribuer.")
                    else:
                        # Construire le planning par jour × sous-agriculteur
                        rows = []
                        for _, day_row in df_plan.groupby("date").agg({"tonnes_jour":"sum"}).reset_index().iterrows():
                            day_total = day_row["tonnes_jour"]
                            for sub_nom, sub_ton in sub_weights.items():
                                share = (sub_ton / total_weight) * day_total
                                # Arrondi à la dizaine supérieure
                                share = math.ceil(share / 10) * 10 if share > 0.5 else 0
                                if share > 0:
                                    rows.append({
                                        "Date":         day_row["date"].strftime("%d/%m"),
                                        "Agriculteur":  sub_nom,
                                        "Tonnes/Jour":  int(share),
                                    })
                        
                        df_internal = pd.DataFrame(rows)
                        if not df_internal.empty:
                            # Pivot
                            piv = df_internal.pivot_table(
                                index="Agriculteur",
                                columns="Date",
                                values="Tonnes/Jour",
                                fill_value=0,
                                aggfunc="sum"
                            )
                            
                            # Ajouter colonne TOTAL
                            piv.insert(0, "TOTAL", piv.sum(axis=1))
                            piv = piv.sort_values("TOTAL", ascending=False)
                            
                            # Ligne TOTAL JOUR
                            totals_row = piv.sum(axis=0)
                            totals_row.name = "── TOTAL JOUR ──"
                            piv = pd.concat([piv, totals_row.to_frame().T])
                            
                            st.dataframe(piv, use_container_width=True,
                                       height=min(800, (len(df_sub)+2) * 35 + 50))
                            
                            # Export
                            st.download_button(
                                f"⬇️ Exporter planning interne {centre_nom}",
                                data=piv.to_csv(index=True).encode("utf-8-sig"),
                                file_name=f"planning_interne_{centre_nom.replace(' ','_')}.csv",
                                mime="text/csv",
                            )
            except Exception as e:
                st.error(f"Erreur lecture planning: {e}")