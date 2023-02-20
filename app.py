import streamlit as st
import pandas as pd

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
form.text_input('Name', placeholder=None, label_visibility="visible")
form.number_input('ID')

submitted = form.form_submit_button('Generate PDF File')
