import streamlit as st
import pandas as pd

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
form.text_input('Name')
form.number_input('ID')
form.number_input('Day Rate (SAR)')
form.text_input('Wellsite Team Leader Name')

submitted = form.form_submit_button('Generate PDF File')
