import streamlit as st
import pandas as pd
from PIL import Image
import tempfile
import os
from datetime import datetime
from io import BytesIO

# Importation des fonctions d'exportation et de génération de codes
# Assurez-vous que ces fichiers (export_pdf.py, export_excel.py, export_to_word_with_page_break.py, generator.py)
# sont présents dans le même répertoire que app.py sur Render.
try:
    from export_pdf import export_to_pdf
    from export_excel import export_to_excel
    from export_to_word_with_page_break import export_to_word
    from generator import generate_qr_code_with_text, generate_code128_with_text, generate_code39_with_text
except ImportError as e:
    st.error(f"Erreur d'importation : Assurez-vous que tous les fichiers (export_pdf.py, export_excel.py, export_to_word_with_page_break.py, generator.py) sont présents. Détails : {e}")
    st.stop() # Arrête l'exécution de l'application si les imports échouent.


st.set_page_config(page_title="Générateur d'Étiquettes Pelichet", layout="centered")

# Logo
# Utilisation de os.path.join et os.path.dirname(__file__) pour une meilleure portabilité
# Assurez-vous que 'LOGO-PELICHET.jpg' est dans le même répertoire que app.py
logo_path = os.path.join(os.path.dirname(__file__), "LOGO-PELICHET.jpg")
if os.path.exists(logo_path):
    st.image(logo_path, width=200)
else:
    st.warning("Le fichier logo 'LOGO-PELICHET.jpg' n'a pas été trouvé. Veuillez vous assurer qu'il est dans le même répertoire.")

st.title("🎯 Générateur d'Étiquettes Pelichet")

# Section de téléchargement de fichier
uploaded_file = st.file_uploader("📁 Sélectionner un fichier Excel avec une colonne 'Code'", type=["xlsx"])

# Options de format de sortie et de type de code
col1, col2 = st.columns(2)
format_output = col1.radio("📄 Format de sortie", ["PDF", "Word", "Excel"])
code_type = col2.selectbox("🔠 Type de code", ["QR", "Code128", "Code39"])

# Options de couleur
col3, col4 = st.columns(2)
code_color = col3.color_picker("🎨 Couleur du code", "#000000")
text_color = col4.color_picker("🎨 Couleur du texte", "#000000")

st.markdown("### 📐 Paramètres d'impression")

# Paramètres de dimensions de l'étiquette
col_a, col_b = st.columns(2)
label_width = col_a.number_input("Largeur étiquette (mm)", 10, 300, 50)
label_height = col_b.number_input("Hauteur étiquette (mm)", 10, 300, 40)

# Paramètres d'espacement
col_c, col_d = st.columns(2)
spacing_x = col_c.number_input("Espacement horizontal (mm)", 0, 50, 0)
spacing_y = col_d.number_input("Espacement vertical (mm)", 0, 50, 0)

# Paramètres de marge
col_e, col_f = st.columns(2)
margin_top = col_e.number_input("Marge haut (mm)", 0, 50, 0)
margin_left = col_f.number_input("Marge gauche (mm)", 0, 50, 0)

# Paramètres de grille
col_g, col_h = st.columns(2)
cols = col_g.number_input("Nombre de colonnes", 1, 10, 3)
rows = col_h.number_input("Nombre de lignes", 1, 30, 7)

# Bouton de génération
if uploaded_file and st.button("🚀 Générer les étiquettes"):
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier Excel. Veuillez vérifier le format. Détails : {e}")
        df = None # S'assure que df est None en cas d'erreur

    if df is not None:
        if "Code" not in df.columns:
            st.error("❌ Le fichier Excel doit contenir une colonne nommée 'Code'.")
        elif df["Code"].empty:
            st.warning("⚠️ La colonne 'Code' est vide. Aucune étiquette ne sera générée.")
        else:
            with st.spinner("Génération des images..."):
                images = []
                temp_files = []
                try:
                    for i, code in enumerate(df["Code"]):
                        img = None
                        # Convertir le code en chaîne de caractères pour éviter les erreurs de type
                        code_str = str(code)
                        if code_type == "QR":
                            # Paramètres QR Code : (code, scale, code_color, text, text_size, text_color, border, label_width, label_height)
                            img = generate_qr_code_with_text(code_str, 10, code_color, code_str, 20, text_color, 5, label_width, label_height)
                        elif code_type == "Code128":
                            # Paramètres Code128 : (code, height, code_color, label_width, label_height)
                            img = generate_code128_with_text(code_str, 50, code_color, label_width, label_height)
                        else: # Code39
                            # Paramètres Code39 : (code, height, code_color, label_width, label_height)
                            img = generate_code39_with_text(code_str, 50, code_color, label_width, label_height)

                        if img:
                            tmp_path = os.path.join(tempfile.gettempdir(), f"etiquette_{i}.png")
                            img.save(tmp_path, format="PNG")
                            images.append(tmp_path)
                            temp_files.append(tmp_path)
                        else:
                            st.warning(f"Impossible de générer l'image pour le code : {code_str}. Le code sera ignoré.")

                    if not images:
                        st.warning("Aucune image n'a pu être générée. Veuillez vérifier vos données et paramètres.")
                    else:
                        now = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"etiquettes_{code_type}_{now}.{ 'pdf' if format_output == 'PDF' else 'docx' if format_output == 'Word' else 'xlsx' }"

                        # Utilisation d'un bloc try-finally pour s'assurer que les fichiers temporaires sont nettoyés
                        # même en cas d'erreur d'exportation.
                        with tempfile.TemporaryDirectory() as tmpdir:
                            outpath = os.path.join(tmpdir, filename)
                            try:
                                if format_output == "PDF":
                                    export_to_pdf(images, outpath, cols, rows, label_width, label_height, spacing_x, spacing_y, margin_top, margin_left)
                                elif format_output == "Excel":
                                    # Pour l'export Excel, nous passons le DataFrame directement.
                                    # Assurez-vous que export_excel.py gère la génération d'images ou l'intégration
                                    # des codes-barres directement dans le fichier Excel sans passer par les images temporaires
                                    # si c'est plus efficace. Sinon, il utilisera les images générées.
                                    export_to_excel(df, outpath, code_type, 50, code_color, label_width, label_height)
                                else: # Word
                                    export_to_word(images, outpath, cols, rows, label_width, label_height, spacing_x, spacing_y, margin_top, margin_left)

                                with open(outpath, "rb") as f:
                                    st.success("✅ Fichier généré avec succès")
                                    st.download_button("📥 Télécharger le fichier", f.read(), file_name=filename, mime="application/octet-stream")
                            except Exception as e:
                                st.error(f"❌ Erreur lors de l'exportation du fichier {format_output}. Détails : {e}")
                            finally:
                                # Nettoyage des fichiers temporaires après l'exportation
                                for f in temp_files:
                                    if os.path.exists(f):
                                        try:
                                            os.remove(f)
                                        except OSError as e:
                                            st.warning(f"Impossible de supprimer le fichier temporaire {f}. Détails : {e}")
                except Exception as e:
                    st.error(f"❌ Une erreur inattendue s'est produite lors de la génération des images ou de l'exportation. Détails : {e}")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
