import re
import random
import pandas as pd
import streamlit as st
from docx.api import Document
import openpyxl
from io import BytesIO

# Liste des variantes de code postal à traiter spécifiquement
CODE_POSTAL_VARIANTS = ['code_postal', 'codepostal', 'codePostal', 'CodePostal']

def clean_zeros_from_empty_cells(df, code_postal_variants=None):
    """
    Remplace les '0' non souhaités (issus de cellules vides) par '' dans le DataFrame,
    sauf pour les colonnes de code postal.
    """
    if code_postal_variants is None:
        code_postal_variants = CODE_POSTAL_VARIANTS
    for col in df.columns:
        # On ne touche pas aux colonnes de code postal
        if col in code_postal_variants or col.lower() in [v.lower() for v in code_postal_variants]:
            continue
        # Remplace '0' (str ou int) par '' si la colonne n'est pas code postal
        df[col] = df[col].replace(['0', 0], '')
    return df

class SpunGenerator:
    def __init__(self):
        self.variable_pattern = r'\$(\w+)'

    def replace_variables(self, text, variables_dict):
        """Remplace les variables avec gestion spécifique des codes postaux et des 0 indésirables"""
        for var, value in variables_dict.items():
            # Gestion spéciale pour les codes postaux
            if var in CODE_POSTAL_VARIANTS or var.lower() in [v.lower() for v in CODE_POSTAL_VARIANTS]:
                if pd.isna(value) or str(value).lower() in ['', 'nan', 'none']:
                    value = ''
                else:
                    try:
                        value = f"{int(float(value)):05d}"  # Format 5 chiffres, retire .0 si float
                    except:
                        value = str(value).zfill(5)
            else:
                # Pour toutes les autres variables : supprime les 0 "vides"
                if pd.isna(value) or str(value).lower() in ['', 'nan', 'none', '0']:
                    value = ''
                else:
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    value = str(value)
            text = text.replace(f'${var}', value)
        return text

    def choose_option(self, options):
        return random.choice(options) if options else ''

    def find_matching_brace(self, text, start):
        count = 1
        pos = start + 1
        while count > 0 and pos < len(text):
            if text[pos] == '{':
                count += 1
            elif text[pos] == '}':
                count -= 1
            pos += 1
        return pos if count == 0 else -1

    def process_simple_options(self, text):
        while True:
            match = re.search(r'{([^{}]+)}', text)
            if not match:
                break
            options = [opt.strip() for opt in match.group(1).split('|')]
            chosen = self.choose_option(options)
            text = text[:match.start()] + chosen + text[match.end():]
        return text

    def process_paragraph_options(self, text):
        def split_options(content):
            options = []
            current = ''
            depth = 0
            for char in content + '|':
                if char == '{':
                    depth += 1
                elif char == '}':
                    depth -= 1
                if char == '|' and depth == 0:
                    if current.strip():
                        options.append(current.strip())
                    current = ''
                else:
                    current += char
            if current.strip():
                options.append(current.strip())
            return options

        result = text
        pattern = r'\{\{([^{}]|{[^{}]*})*\}\}'
        while True:
            match = re.search(pattern, result, re.DOTALL)
            if not match:
                break
            full_match = match.group(0)
            content = full_match[2:-2]
            options = split_options(content)
            if options:
                chosen = self.choose_option(options)
                processed = self.process_simple_options(chosen)
                result = result[:match.start()] + processed + result[match.end():]
            else:
                result = result[:match.start()] + content + result[match.end():]
        return result

    def generate_spun(self, text, variables_dict):
        text = self.process_paragraph_options(text)
        text = self.process_simple_options(text)
        text = self.replace_variables(text, variables_dict)
        return text

def process_input_file(file_bytes):
    try:
        if file_bytes.name.endswith('.docx'):
            doc = Document(BytesIO(file_bytes.read()))
            return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
        else:
            content = file_bytes.read()
            if isinstance(content, bytes):
                return content.decode('utf-8')
            return content
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier: {str(e)}")
        raise

def generate_spuns(input_text, df_variables, num_spuns):
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
    text_file = st.file_uploader("Fichier texte (.txt ou .docx)", type=['txt', 'docx'])
    excel_file = st.file_uploader("Fichier Excel des variables", type=['xlsx'])
    num_spuns = st.number_input("Nombre de spuns à générer", min_value=1, value=1)
    preview_count = st.number_input("Nombre de spuns à prévisualiser", min_value=1, max_value=5, value=1)
    if st.button("Générer les spuns") and text_file and excel_file:
        try:
            with st.spinner('Génération des spuns en cours...'):
                # Lecture du fichier avec gestion des codes postaux en texte
                df_variables = pd.read_excel(
                    excel_file,
                    dtype={v: str for v in CODE_POSTAL_VARIANTS},
                    keep_default_na=False
                )
                # Nettoyage des 0 indésirables
                df_variables = clean_zeros_from_empty_cells(df_variables)
                input_text = process_input_file(text_file)
                df_results = generate_spuns(input_text, df_variables, num_spuns)
                st.subheader("Prévisualisation des spuns générés")
                for i in range(min(preview_count, len(df_results))):
                    with st.expander(f"Spun #{df_results.iloc[i]['Spun_ID']}", expanded=i == 0):
                        st.text_area(
                            "Texte généré",
                            value=df_results.iloc[i]['Texte_Généré'],
                            height=200,
                            disabled=True
                        )
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
