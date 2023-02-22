import streamlit as st
import pandas as pd
from openpyxl import load_workbook
# from asposecells.api import Workbook
from pdfrw import PdfWriter
from io import BytesIO, StringIO
from xlsx2html import xlsx2html
from xhtml2pdf import pisa
import datetime
import calendar

st.title('BakerTimeSheetGenerator')
form = st.form('input_form')
employee_name = form.text_input('Name')
employee_id = form.number_input('ID', step=1)
employee_rate = form.number_input('Day Rate (SAR)')
date_in = form.date_input('Date In')
date_out = form.date_input('Date Out')
wstl_name = form.selectbox('Wellsite Team Leader', ['KL', 'PO', 'SB', 'AM'])
# expander = form.expander('Other settings')
# expander.write('ikhy')

submitted = form.form_submit_button('Generate PDF File')

months_dict = {
    1: 'JAN',
    2: 'FEB',
}

if submitted:
    with st.spinner('Working on your timesheet...'):

        st.write(date_in, date_out)
        #date_in_stripped = datetime.datetime.strptime(date_in, "%Y-%m-%d")
        st.write(date_in.month, calendar.month_abbr[date_in.month] + date_in.year)


        output = BytesIO()
        wb = load_workbook(filename=r'template.xlsx', read_only=False)
        ws = wb['timesheet']
        ws['Q2']= employee_name
        ws['Q3']= employee_id
        ws['Q5']= calendar.month_abbr[date_in.month]
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
        #result_file = open(pdf_output, "w+b")
        pisa_stat = pisa.CreatePDF(out_stream.getvalue(), dest=pdf_output)
        #pdf_output.close()
        #pdf_output.seek(0)

        st.download_button(
            label="Download pdf Page",
            data=pdf_output.getvalue(),
            file_name="report.pdf",
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

