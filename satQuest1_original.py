import pandas as pd

# Charger le fichier Excel
fichier_entree = "C:/Users/kossi.dovenon/Downloads/Questionnaire_initial/resultats_quiz_QUESTIONNAIRE INITIAL_2025-06-01_20-36.xlsx"  # Remplace avec le chemin de ton fichier
df = pd.read_excel(fichier_entree)

# Vérifier les colonnes présentes
print("Colonnes détectées :", df.columns)

# Renommer les colonnes si besoin (adapter en fonction des noms réels)
df.columns = ["Nom", "Question", "Réponse", "Autre"]

# Transformer le fichier pour que les questions deviennent des colonnes
df_pivot = df.pivot_table(index="Nom", columns="Question", values="Réponse", aggfunc="first")

# Réinitialiser l'index pour obtenir un fichier propre
df_pivot.reset_index(inplace=True)

# Sauvegarder dans un nouveau fichier Excel
fichier_sortie = "C:/Users/kossi.dovenon/Downloads/Questionnaire_initial/resultats_quiz_QUESTIONNAIRE INITIAL_2025-06-01_20-36fichier_reorganise.xlsx"
df_pivot.to_excel(fichier_sortie, index=False)

print(f"Fichier transformé et enregistré sous {fichier_sortie}")