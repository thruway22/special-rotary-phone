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
date_start = form.date_input('Date Start')
date_end = form.date_input('Date End')
rig_name = form.text_input('Rig Name')
wstl_list = ['Ken Lynn', 'Pete Riley', 'Steve Baranyi', 'Ahmed Mansour']
wstl_name = form.selectbox('Wellsite Team Leader', wstl_list.sort())
# expander = form.expander('Other settings')
# expander.write('ikhy')

submitted = form.form_submit_button('Generate PDF File')

if submitted:
    if date_start > date_end:
        st.error('Starting date is later than ending date.')

    else:
        with st.spinner('Working on your timesheet...'):
            output = BytesIO()
            wb = load_workbook(filename=r'template.xlsx', read_only=False)
            ws = wb['timesheet']

            month_start = 1
            month_end = calendar.monthrange(date_start.year, date_start.month)[1] + 1
            for day in range(month_start, month_end):
                cell_a = 'A' + str(day + 1)
                ws[cell_a] = day

            shift_start = date_start.day
            shift_end = month_end if date_end.month > date_start.month else date_end.day + 1
            for shift in range(shift_start, shift_end):
                cell_b = 'B' + str(shift + 1)
                cell_d = 'D' + str(shift + 1)
                ws[cell_b] = 'ARAMCO'
                ws[cell_d] = rig_name

                    
            ws['Q2']= employee_name
            ws['Q3']= employee_id
            ws['Q5']= str(calendar.month_abbr[date_start.month].upper()) + ' ' + str(date_start.year) # month year

            hitch = len(range(shift_start, shift_end)) # total shift days
            ws['O8']= hitch
            ws['Q8']= employee_rate
            ws['T8']= hitch * employee_rate
            ws['T19']= hitch * employee_rate

            ws['O22']= employee_name
            ws['O24']= wstl_name
            
            wb.copy_worksheet(ws)
            ws2 = wb['timesheet Copy']
            ws2.title = "timesheet 2"
            wb.save(output)

            st.download_button(
                label="Download Excel workbook",
                data=output.getvalue(),
                file_name="workbook",
                mime="application/vnd.ms-excel"
            )

            # output_html = BytesIO()
            # html = xlsx2html(output, b'output.html')
            #####################################
            # out_stream = xlsx2html(output)
            # out_stream.seek(0)
            

            # st.download_button(
            #     label="Download Html Page",
            #     data=out_stream.read(),
            #     file_name="report.html",
            #     mime="application/octet-stream"
            # )
            # ####################################
            # pdf_output = BytesIO()
            # #result_file = open(pdf_output, "w+b")
            # pisa_stat = pisa.CreatePDF(out_stream.getvalue(), dest=pdf_output)
            # #pdf_output.close()
            # #pdf_output.seek(0)

            # st.download_button(
            #     label="Download pdf Page",
            #     data=pdf_output.getvalue(),
            #     file_name="report.pdf",
            #     mime="application/octet-stream")

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

