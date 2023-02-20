import streamlit as st
import pandas as pd

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
form.text_input('Name')
form.number_input('ID')
form.number_input('Day Rate (SAR)')
form.selectbox('Wellsite Team Leader', ['KL', 'PO', 'SB', 'AM'])

submitted = form.form_submit_button('Generate PDF File')
