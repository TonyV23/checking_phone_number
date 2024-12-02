from src.compare import Comparateur

def main():
    # Chemins des fichiers
    fichier_excel = "data/input/line_to_remove_obr.xls"
    fichier_sql = "data/input/exclude_invoices.sql"
    fichier_sortie_communs = "data/output/numeros_communs.xlsx"
    fichier_sortie_uniques = "data/output/numeros_unique_excel.xlsx"
    
    try:
        # Création de l'instance du comparateur
        comparateur = Comparateur(fichier_excel, fichier_sql)
        
        # Comparaison des numéros
        resultats = comparateur.comparer_numeros()
        
        # Affichage des résultats
        print(f"Nombre de numéros communs trouvés : {len(resultats['communs'])}")
        print(f"Nombre de numéros uniquement dans Excel : {len(resultats['unique_excel'])}")
        
        # Sauvegarde des résultats
        comparateur.sauvegarder_resultats(
            resultats,
            fichier_sortie_communs,
            fichier_sortie_uniques
        )
        
    except Exception as e:
        print(f"Une erreur s'est produite : {str(e)}")

if __name__ == "__main__":
    main()