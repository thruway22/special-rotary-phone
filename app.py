import streamlit as st
import pandas as pd
from openpyxl import load_workbook
# from asposecells.api import Workbook
from pdfrw import PdfWriter
from io import BytesIO
from xlsx2html import xlsx2html
from xhtml2pdf import pisa

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
        output = BytesIO()
        wb = load_workbook(filename=r'template.xlsx', read_only=False)
        ws = wb['timesheet']
        ws['Q2']= "Saleh"
        ws['A2']=10
        wb.save(output)

        st.download_button(
            label="Download Excel workbook",
            data=output.getvalue(),
            file_name="workbook",
            mime="application/vnd.ms-excel"
        )

        # output_html = BytesIO()
        # html = xlsx2html(output, b'output.html')
        
        out_stream = xlsx2html(output)
        out_stream.seek(0)
        

        st.download_button(
            label="Download Html Page",
            data=out_stream.read(),
            file_name="report.html",
            mime="application/octet-stream"
        )
        pdf_output = BytesIO()
        result_file = open(b'test.pdf', "w+b")
        pdf_out = pisa.CreatePDF(out_stream.read(), dest=result_file)
        result_file.close()

        st.download_button(
            label="Download Html Page",
            data=result_file,
            file_name="report.html",
            mime="application/octet-stream")

        # pdf_out = BytesIO()
        # y = PdfWriter()
        # #y.addpage()
        # y.write(pdf_out)
        # # ws_range = ws.iter_rows()
        # # for row in ws_range:
        # #     s = ''
        # #     for cell in row:
        # #         if cell.value is None:
        # #             s = s
        # #         else:
        # #             s += str(cell.value) #.rjust(10) + ' '
        # #     pw.writeLine(s)
        # # pw.savePage()
        # # pw.close()

        # st.download_button(label="Export_Report",
        #             data=pdf_out.getvalue(),
        #             file_name="test.pdf",
        #             mime='application/octet-stream')

        st.write('test')

