import pandas as pd
import streamlit as st

st.set_page_config(page_title="PAF Prévis", page_icon="🛂", layout="centered", initial_sidebar_state="auto")


st.title('🛂 PAF Prévis') 
st.divider()


st.sidebar.info("")


hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

