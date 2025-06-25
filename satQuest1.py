import pandas as pd

# Charger le fichier Excel
fichier_entree = "C:/Users/kossi.dovenon/Desktop/traitementFichier/passage_des_questionnaires_du_16_06_2025.xlsx"  # Remplace avec le chemin de ton fichier
df = pd.read_excel(fichier_entree)

# Vérifier les colonnes présentes
print("Colonnes détectées :", df.columns)

# Renommer les colonnes si besoin
df.columns = ["Nom", "Question", "Réponse", "Commentaire"]

# Créer un pivot des réponses
df_pivot = df.pivot_table(index="Nom", columns="Question", values="Réponse", aggfunc="first")

# Extraire les commentaires par utilisateur (en prenant le premier si plusieurs)
df_commentaires = df.groupby("Nom")["Commentaire"].first().reset_index()

# Fusionner le tableau pivoté avec les commentaires
df_final = df_pivot.reset_index().merge(df_commentaires, on="Nom", how="left")

# Sauvegarder
fichier_sortie = "C:/Users/kossi.dovenon/Desktop/traitementFichier/passage_des_questionnaires_du_16_06_2025_fichier_reorganise.xlsx"
df_final.to_excel(fichier_sortie, index=False)

print(f"Fichier transformé et enregistré sous {fichier_sortie}")
