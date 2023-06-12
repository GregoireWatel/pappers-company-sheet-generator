import streamlit as st
import openai

def set_sidebar():
    for i in range(28): st.sidebar.text('')
    st.sidebar.info("GitHub Repository: <https://github.com/ceptln/smart-sales-assistant> \n\n Contact: camille.epitalon@polytechnique.edu")

def set_page_configuration(initial_sidebar_state="expanded"):
    return st.set_page_config(
            page_title="Company Sheet Generator",
            page_icon="🏢",
            # layout="wide",
            initial_sidebar_state=initial_sidebar_state,
            menu_items={
                'Get Help': 'https://github.com/ceptln',
                'Report a bug': "https://github.com/ceptln",
                'About': "This app was built an deployed by Camille Goat Epitalon"
            }
        )

def open_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as infile:
        return infile.read()
    

