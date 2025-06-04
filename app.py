import re
import random
import pandas as pd
import streamlit as st
from docx.api import Document
import openpyxl
from io import BytesIO

class SpunGenerator:
    def __init__(self):
        self.variable_pattern = r'\$(\w+)'

    def replace_variables(self, text, variables_dict):
        """Remplace les variables avec gestion spécifique des codes postaux"""
        for var, value in variables_dict.items():
            # Gestion spéciale pour les codes postaux
            if var.lower() in ['codepostal', 'code_postal']:
                if pd.isna(value) or value in ['', 'nan']:
                    value = ''
                else:
                    # Conversion float => int et formatage sur 5 chiffres
                    try:
                        value = f"{int(float(value)):05d}"
                    except:
                        value = str(value).zfill(5)
            else:
                # Comportement original pour les autres variables
                if pd.isna(value) or value == '':
                    value = ''
                else:
                    value = str(value)
            
            if var in text:
                text = text.replace(f'${var}', value)
        
        return text

    # ... (Les autres méthodes de la classe restent identiques à la version initiale) ...

def generate_spuns(input_text, df_variables, num_spuns):
    """Génère les spuns et retourne un DataFrame"""
    generator = SpunGenerator()
    results = []
    
    for index, row in df_variables.iterrows():
        if index >= num_spuns:
            break
        
        variables_dict = row.to_dict()
        spun_text = generator.generate_spun(input_text, variables_dict)
        spun_text = spun_text.replace('###devider###', '###devider###\n')
        results.append([index + 1, spun_text])
    
    return pd.DataFrame(results, columns=['Spun_ID', 'Texte_Généré'])

def create_streamlit_app():
    st.title("Générateur de Spuns")
    
    # Upload des fichiers
    text_file = st.file_uploader("Fichier texte (.txt ou .docx)", type=['txt', 'docx'])
    excel_file = st.file_uploader("Fichier Excel des variables", type=['xlsx'])
    
    # Nombre de spuns à générer
    num_spuns = st.number_input("Nombre de spuns à générer", min_value=1, value=1)
    
    # Prévisualisation
    preview_count = st.number_input("Nombre de spuns à prévisualiser", min_value=1, max_value=5, value=1)
    
    if st.button("Générer les spuns") and text_file and excel_file:
        try:
            with st.spinner('Génération des spuns en cours...'):
                # Lecture des fichiers avec gestion des codes postaux en texte
                input_text = process_input_file(text_file)
                df_variables = pd.read_excel(
                    excel_file, 
                    dtype={'code_postal': str, 'codepostal': str}  # Force le texte pour les codes postaux
                )
                
                # Génération des spuns
                df_results = generate_spuns(input_text, df_variables, num_spuns)
                
                # Affichage de la prévisualisation
                st.subheader("Prévisualisation des spuns générés")
                for i in range(min(preview_count, len(df_results))):
                    with st.expander(f"Spun #{df_results.iloc[i]['Spun_ID']}", expanded=i==0):
                        st.text_area(
                            "Texte généré",
                            value=df_results.iloc[i]['Texte_Généré'],
                            height=200,
                            disabled=True
                        )
                
                # Téléchargement
                output = BytesIO()
                df_results.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                
                st.download_button(
                    label="Télécharger tous les spuns générés",
                    data=output,
                    file_name="spuns_generes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        except Exception as e:
            st.error(f"Une erreur s'est produite: {str(e)}")

# ... (Les fonctions process_input_file() et le reste du code restent identiques à la version initiale) ...

if __name__ == "__main__":
    create_streamlit_app()
