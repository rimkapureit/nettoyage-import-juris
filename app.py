import streamlit as st
import pandas as pd

st.title("📊 Préparation import pour Juris")

uploaded_file = st.file_uploader("Choisir le fichier Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file, sheet_name="Employes actifs de l'AJQ", skiprows=2)
    st.write("✅ Fichier chargé")

    regions = df["Ds Region"].dropna().unique()
    choix = st.selectbox("Choisir une région", ["Toutes"] + list(regions))

    # -------------------------
    # Fonctions
    # -------------------------

    def normaliser_telephone(x):
        if pd.isna(x):
            return x
        s = "".join(ch for ch in str(x) if ch.isdigit())
        if len(s) >= 10:
            s = s[-10:]
            return f"{s[:3]}-{s[3:6]}-{s[6:]}"
        return x

    def fusionner_tel_poste(row):
        tel = row.get("Téléphone")
        poste = row.get("Poste")

        if pd.isna(tel):
            return tel

        if not pd.isna(poste) and str(poste).strip() != "":
            poste_str = str(poste).strip()
            if poste_str.endswith(".0"):
                poste_str = poste_str[:-2]
            return f"{tel} poste {poste_str}"

        return tel

    # -------------------------
    # Bouton
    # -------------------------

    if st.button("🚀 Lancer le traitement"):

        df_clean = df.copy()

        # Filtre région
        if choix != "Toutes":
            df_clean = df_clean[df_clean["Ds Region"] == choix]

        # Nettoyage colonnes
        df_clean = df_clean.drop(columns=["Column5", "Column10"], errors="ignore")

        # Téléphone
        if "Téléphone" in df_clean.columns:
            df_clean["Téléphone"] = df_clean["Téléphone"].apply(normaliser_telephone)

        # Fusion poste
        if "Téléphone" in df_clean.columns and "Poste" in df_clean.columns:
            df_clean["Téléphone"] = df_clean.apply(fusionner_tel_poste, axis=1)
            df_clean = df_clean.drop(columns="Poste")

        # Nettoyage Fonction
        if "Fonction" in df_clean.columns:
            df_clean["Fonction"] = (
                df_clean["Fonction"].astype(str)
                .str.replace("avocate", "avocat", regex=False)
                .str.replace("employée de soutien", "", regex=False)
                .str.replace("employé de soutien", "", regex=False)
                .str.replace("secrétaire", "", regex=False)
                .str.replace("directeur", "avocat", regex=False)
                .str.replace("directrice", "avocat", regex=False)
                .str.replace("professionel", "", regex=False)
                .str.replace("professionnel", "", regex=False)
                .str.replace("exécutif", "avocat", regex=False)
                .str.replace("avocat stagiaire", "", regex=False)
            )

        # -------------------------
        # Colonnes
        # -------------------------

        colonnes_apres_prenom = [
            "DateNaiss", "Sexe", "Langue", "NAS", "NoBande", "Nsr", "SED"
        ]

        colonnes_fin = [
            "Profil", "AccesDossier", "AccesActivite", "AccesFeuilleTemps",
            "AccesClavardage", "AccesRapStat", "AccesEmploye", "AccesBureau",
            "PROPRIÉTÉ SYSTÈME", "Titre",
            "Numéro d'impliqué permanent",
            "Destinataire par défaut des activités",
            "Calendrier des activités (heure de début)",
            "Calendrier des activités (heure de fin)",
            "Activités (durées des activités en minutes)",
            "Heure de début des activités",
            "Bureaux additionnels"
        ]

        # -------------------------
# Préremplissage SaaS
# -------------------------

valeurs_defaut = {
    "Sexe": "F",
    "Langue": "QC",
    "Profil": "Employé(e)Avocat/Soutien",
    "AccesDossier": "tous",
    "AccesActivite": "tous",
    "AccesFeuilleTemps": "usager courant",
    "AccesClavardage": "usager courant",
    "AccesRapStat": "usager courant",
    "AccesEmploye": "usager courant",
    "AccesBureau": "Sélection",
    "PROPRIÉTÉ SYSTÈME": "Usager Courant",
    "Calendrier des activités (heure de début)": "08:00:00",
    "Calendrier des activités (heure de fin)": "23:30:00",
    "Activités (durées des activités en minutes)": "15",
    "Heure de début des activités": "08:00:00"
}

# Appliquer seulement si vide
for col, val in valeurs_defaut.items():
    if col in df_clean.columns:
        df_clean[col] = df_clean[col].replace("", val)

        # ✅ Ajouter colonnes AVANT réorganisation
        for col in colonnes_apres_prenom + colonnes_fin:
            if col not in df_clean.columns:
                df_clean[col] = ""

        # -------------------------
        # Ordre final
        # -------------------------

        colonnes_base = [
            "Ds Region", "Nom Bureau", "Nom", "Prénom"
        ]

        colonnes_suite = [
            "Fonction", "Code Avocat", "Téléphone",
            "Courriel", "Telecopieur", "Adresse", "Ville", "Code Postal"
        ]

        ordre_final = (
            colonnes_base +
            colonnes_apres_prenom +
            colonnes_suite +
            colonnes_fin
        )

        df_clean = df_clean[[c for c in ordre_final if c in df_clean.columns]]

        # Nettoyage final
        df_clean = df_clean.fillna("")
        df_clean = df_clean.astype(str)

        # -------------------------
        # Affichage
        # -------------------------

        st.dataframe(df_clean)

        # -------------------------
        # Export
        # -------------------------

        output_file = "resultat.xlsx"
        df_clean.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=f,
                file_name="Employes_nettoyes.xlsx"
            )
