import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import asposecells

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
employee_name = form.text_input('Name')
employee_id = form.number_input('ID')
employee_rate = form.number_input('Day Rate (SAR)')
date_in = form.date_input('Date In')
date_out = form.date_input('Date Out')
wstl_name = form.selectbox('Wellsite Team Leader', ['KL', 'PO', 'SB', 'AM'])
# expander = form.expander('Other settings')
# expander.write('ikhy')

submitted = form.form_submit_button('Generate PDF File')

if submitted:
    with st.spinner('Working on your timesheet...'):
        wb = load_workbook(filename=r'template.xlsx', read_only=False)
        ws = wb['timesheet']
        ws['Q2']= "Saleh"
        ouput_excel = wb.save("sample.xlsx")
        output_pdf = Workbook(ouput_excel)
        pdf = ouput_excel.save("timesheet.pdf", SaveFormat.PDF)
        btn = st.download_button(
            label="Download image",
            data=pdf,
            file_name="flower.png",
            #mime="image/png",
        )
        st.write('test')

