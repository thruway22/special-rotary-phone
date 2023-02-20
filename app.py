import streamlit as st
import pandas as pd

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
employee_name = form.text_input('Name')
employee_id = form.number_input('ID')
employee_rate = form.number_input('Day Rate (SAR)')
date_in = form.date_input('Date In')
date_out = form.date_input('Date Out')
wstl_name = form.selectbox('Wellsite Team Leader', ['KL', 'PO', 'SB', 'AM'])
form.expander('Other settings')
expander.write(\"\"\"
    The chart above shows some numbers I picked for you.
    I rolled actual dice for these, so they're *guaranteed* to
    be random.
\"\"\")

submitted = form.form_submit_button('Generate PDF File')
