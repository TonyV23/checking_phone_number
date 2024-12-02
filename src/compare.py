import pandas as pd
import re

class Comparateur:
    def __init__(self, excel_path, sql_path):
        self.excel_path = excel_path
        self.sql_path = sql_path
        
    def lire_excel(self):
        try:
            df = pd.read_excel(self.excel_path)
            return df['ACC_NBR'].astype(str).str.strip().tolist()
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier Excel: {str(e)}")
            return []
            
    def lire_sql(self):
        try:
            with open(self.sql_path, 'r', encoding='utf-8') as file:
                content = file.read()
            
            pattern = r"'(\d+)'"
            numeros = re.findall(pattern, content)
            return [num.strip() for num in numeros]
            
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier SQL: {str(e)}")
            return []
    
    def comparer_numeros(self):
        numeros_excel = self.lire_excel()
        numeros_sql = self.lire_sql()
        
        # Convertir en ensembles pour les opérations d'ensemble
        set_excel = set(numeros_excel)
        set_sql = set(numeros_sql)
        
        # Trouver les numéros communs
        numeros_communs = set_excel & set_sql
        
        # Trouver les numéros uniquement dans Excel
        numeros_unique_excel = set_excel - set_sql
        
        return {
            'communs': list(numeros_communs),
            'unique_excel': list(numeros_unique_excel)
        }
    
    def sauvegarder_resultats(self, resultats, chemin_sortie_communs, chemin_sortie_uniques):
        try:
            # Sauvegarde des numéros communs
            df_communs = pd.DataFrame(resultats['communs'], columns=['Numeros_Communs'])
            df_communs.to_excel(chemin_sortie_communs, index=False)
            print(f"Numéros communs sauvegardés dans {chemin_sortie_communs}")
            
            # Sauvegarde des numéros uniques à Excel
            df_uniques = pd.DataFrame(resultats['unique_excel'], columns=['Numeros_Unique_Excel'])
            df_uniques.to_excel(chemin_sortie_uniques, index=False)
            print(f"Numéros uniques Excel sauvegardés dans {chemin_sortie_uniques}")
            
        except Exception as e:
            print(f"Erreur lors de la sauvegarde des résultats: {str(e)}")