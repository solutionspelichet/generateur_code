
# Code de secours : message de test seulement pour confirmer l'affichage
import streamlit as st

def main():
    st.set_page_config(page_title="Test Interface", layout="centered")
    st.title("🧾 Interface OK")
    st.success("✅ Ceci est une version de secours. Le code fonctionne.")

if __name__ == "__main__":
    main()
