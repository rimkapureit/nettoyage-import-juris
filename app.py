import streamlit as st
import pandas as pd

st.title("📊 Nettoyage des employés AJQ")

# Upload fichier
uploaded_file = st.file_uploader("Choisir le fichier Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file, sheet_name="Employes actifs de l'AJQ", skiprows=2)

    st.write("✅ Fichier chargé")

    # Liste des régions
    regions = df["Ds Region"].dropna().unique()

    choix = st.selectbox("Choisir une région", ["Toutes"] + list(regions))

    # Fonctions
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

    # Bouton d'exécution
    if st.button("🚀 Lancer le traitement"):

        df_clean = df.copy()

        # Filtrer région
        if choix != "Toutes":
            df_clean = df_clean[df_clean["Ds Region"] == choix]

        # Nettoyage colonnes
        df_clean = df_clean.drop(columns=["Column5", "Column10"], errors="ignore")

        # Téléphone
        if "Téléphone" in df_clean.columns:
            df_clean["Téléphone"] = df_clean["Téléphone"].apply(normaliser_telephone)

        # Fusion Poste
        if "Téléphone" in df_clean.columns and "Poste" in df_clean.columns:
            df_clean["Téléphone"] = df_clean.apply(fusionner_tel_poste, axis=1)
            df_clean = df_clean.drop(columns="Poste")

        # Fonction
        if "Fonction" in df_clean.columns:
            df_clean["Fonction"] = (
                df_clean["Fonction"].astype(str)
                .str.replace("avocate", "avocat", regex=False)
                .str.replace("employée de soutien", "", regex=False)
                .str.replace("employé de soutien", "avocat", regex=False)
                .str.replace("secrétaire", "", regex=False)
                .str.replace("directeur", "avocat", regex=False)
                .str.replace("directrice", "avocat", regex=False)
                .str.replace("professionel", "avocat", regex=False)
                .str.replace("professionnel", "avocat", regex=False)
                .str.replace("professionnelle", "avocat", regex=False)
                .str.replace("exécutif", "avocat", regex=False)
            )

        # Ordre colonnes
        wanted_order = [
            "Ds Region", "Nom Bureau", "Nom", "Prénom", "Fonction", "Code Avocat",
            "Téléphone", "Courriel", "Telecopieur", "Adresse", "Ville", "Code Postal"
        ]

        df_clean = df_clean[[c for c in wanted_order if c in df_clean.columns]]

        # Affichage
        st.dataframe(df_clean)

        # Export
        output_file = "resultat.xlsx"
        df_clean.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=f,
                file_name="Employes_nettoyes.xlsx"
            )