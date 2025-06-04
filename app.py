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
        """Remplace les variables avec gestion avancée des codes postaux"""
        for var, value in variables_dict.items():
            # Gestion spéciale pour toutes les variantes de code postal
            if var.lower() in ['codepostal', 'code_postal', 'codeposte']:
                if pd.isna(value) or value == '':
                    value = ''
                else:
                    # Conversion en texte brut avec zéros initiaux
                    if isinstance(value, float):
                        value = f"{int(value):05d}"  # Formatage sur 5 chiffres
                    else:
                        value = str(value).zfill(5)
            else:
                # Comportement natif pour les autres variables
                if pd.isna(value) or value == '':
                    value = ''
                else:
                    # Conversion float => int si nécessaire
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    value = str(value)
            text = text.replace(f'${var}$', value)
        return text

    # ... (le reste des méthodes de la classe reste inchangé) ...

def generate_spuns(input_text, df_variables, num_spuns):
    """Génère les spuns avec gestion des types de colonnes"""
    generator = SpunGenerator()
    results = []
    
    for index, row in df_variables.iterrows():
        if index >= num_spuns:
            break
        
        # Conversion manuelle pour compatibilité
        variables_dict = {}
        for col in df_variables.columns:
            col_lower = col.lower()
            if col_lower in ['codepostal', 'code_postal', 'codeposte']:
                variables_dict[col] = str(row[col]).zfill(5) if not pd.isna(row[col]) else ''
            else:
                val = row[col]
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                variables_dict[col] = '' if pd.isna(val) else str(val)
        
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
                # Lecture du fichier avec gestion des codes postaux
                df_variables = pd.read_excel(
                    excel_file,
                    dtype={'CodePostal': str, 'code_postal': str, 'codepostal': str},
                    keep_default_na=False
                )
                
                # Génération des spuns
                input_text = process_input_file(text_file)
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
                
                # Téléchargements
                csv_data = df_results.to_csv(index=False).encode('utf-8')
                excel_buffer = BytesIO()
                df_results.to_excel(excel_buffer, index=False, engine='openpyxl')
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        "📥 Télécharger en CSV",
                        data=csv_data,
                        file_name="spuns.csv",
                        mime="text/csv"
                    )
                with col2:
                    st.download_button(
                        "📊 Télécharger en Excel",
                        data=excel_buffer.getvalue(),
                        file_name="spuns.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        except Exception as e:
            st.error(f"Erreur critique : {str(e)}")

if __name__ == "__main__":
    create_streamlit_app()
