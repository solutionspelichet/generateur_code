import streamlit as st
import traceback

def main():
    st.set_page_config(page_title="Test Interface Streamlit", layout="centered")
    st.title("🧾 Générateur de codes (TEST)")
    st.success("✅ L'interface Streamlit est bien chargée.")

    try:
        code_type = st.selectbox("Type de code", ["QR Code", "Code 39", "Code 128"])
        font_size = st.slider("Taille du texte", 6, 36, 12)
        label_w = st.number_input("Largeur image (mm)", value=50.0)
        label_h = st.number_input("Hauteur image (mm)", value=25.0)
        cols = st.number_input("Colonnes", min_value=1, max_value=10, value=2)
        rows = st.number_input("Lignes", min_value=1, max_value=30, value=6)
        spacing_x = st.number_input("Espacement horizontal (mm)", value=5.0)
        spacing_y = st.number_input("Espacement vertical (mm)", value=5.0)
        margin_top = st.number_input("Marge haute (mm)", value=10.0)
        margin_left = st.number_input("Marge gauche (mm)", value=10.0)
        st.file_uploader("Fichier Excel", type=["xlsx"])
        st.button("Générer")

    except Exception as e:
        st.error("💥 Une erreur est survenue dans le code :")
        st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
