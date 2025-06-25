import pandas as pd

# Charger toutes les feuilles
fichier_entree = "C:/Users/kossi.dovenon/Desktop/traitementFichier/passage_des_questionnaires_du_14_06_2025.xlsx"
feuilles = pd.read_excel(fichier_entree, sheet_name=None)  # Toutes les feuilles

# Fichier de sortie unique
fichier_sortie = "C:/Users/kossi.dovenon/Desktop/traitementFichier/passage_des_questionnaires_du_14_06_2025_fichier_reorganise.xlsx"

# Écrire dans un seul fichier avec plusieurs onglets
with pd.ExcelWriter(fichier_sortie, engine='openpyxl') as writer:
    for nom_feuille, df in feuilles.items():
        print(f"Traitement de la feuille : {nom_feuille}")

        # Vérifier les colonnes
        print("Colonnes détectées :", df.columns)

        # Renommer les colonnes si besoin
        df.columns = ["Nom", "Question", "Réponse", "Commentaire"]

        # Créer le pivot
        df_pivot = df.pivot_table(index="Nom", columns="Question", values="Réponse", aggfunc="first")

        # Extraire les commentaires
        df_commentaires = df.groupby("Nom")["Commentaire"].first().reset_index()

        # Fusion finale
        df_final = df_pivot.reset_index().merge(df_commentaires, on="Nom", how="left")

        # Écrire dans un onglet du fichier final
        df_final.to_excel(writer, sheet_name=nom_feuille, index=False)

print(f"Fichier unique généré avec toutes les feuilles : {fichier_sortie}")
