from tkinter import font

from reportlab.pdfgen import canvas
from reportlab.platypus import (SimpleDocTemplate, Paragraph, PageBreak, Image, Spacer, Table, TableStyle)
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.pagesizes import mm  # LETTER, inch
from reportlab.rl_config import defaultPageSize
from reportlab.lib.pagesizes import landscape
from reportlab.graphics.shapes import Line, LineShape, Drawing
from reportlab.lib.colors import Color
from reportlab.lib import utils, colors
import pandas as pd
from datetime import date
import os
import errno


################ create folder to save prepared reports
current_date_2 = date.today().strftime("%Y-%m-%d")  # create date
file_date = ('File Date: ' + current_date_2)
folder_name = (current_date_2)
folder_path = os.path.join('./' + folder_name + '/')
try:
    os.makedirs(folder_name)
except OSError as e:
    if e.errno != errno.EEXIST:
        raise
    pass

################ import and manipulate files

## report_A
# preparation
report_A = pd.read_excel('./Data/Financial Sample.xlsx')
report_A.drop_duplicates(subset='Country').dropna(subset=['Segment']).sort_values(by=['Month Name'])
report_A = report_A[['Segment', 'Country']].rename(columns={'Segment': 's', 'Country': 'c'})
# save in excel
report_A_save_name = os.path.join(folder_path + 'report_A ' + current_date_2 + '.xlsx')
report_A.to_excel(report_A_save_name, index=False)

## report_B
report_B = pd.read_excel('./Data/Financial Sample.xlsx')
report_B.drop_duplicates(subset='Country').dropna(subset=['Segment']).sort_values(by=['Month Name'])
# save in excel
report_B_save_name = os.path.join(folder_path + 'report_B ' + current_date_2 + '.xlsx')
report_B.to_excel(report_B_save_name, index=False)

## report_C
report_C = pd.read_excel('./Data/Financial Sample.xlsx')
report_C.drop_duplicates(subset='Country').dropna(subset=['Segment']).sort_values(by=['Month Name'])
# save in excel
report_C_save_name = os.path.join(folder_path + 'report_C ' + current_date_2 + '.xlsx')
report_C.to_excel(report_C_save_name, index=False)

## report_D
report_D = pd.read_excel('./Data/Financial Sample.xlsx')
report_D.drop_duplicates(subset='Country').dropna(subset=['Segment']).sort_values(by=['Month Name'])
# save in excel
report_D_save_name = os.path.join(folder_path + 'report_D ' + current_date_2 + '.xlsx')
report_D.to_excel(report_D_save_name, index=False)

## save all reports in one excel
report_all_in_one_save_name = os.path.join(folder_path + 'report_all_in_one ' + current_date_2 + '.xlsx')

with pd.ExcelWriter(report_all_in_one_save_name) as writer:
    report_A.to_excel(writer, index=False, sheet_name='report_A')
    report_B.to_excel(writer, index=False, sheet_name='report_B')
    report_C.to_excel(writer, index=False, sheet_name='report_C')
    report_D.to_excel(writer, index=False, sheet_name='report_D')


################ functions

# resize image
def get_image(path, width=1 * mm):
    img = utils.ImageReader(path)
    iw, ih = img.getSize()
    aspect = ih / float(iw)
    return Image(path, width=width, height=(width * aspect))


# get current date
current_date_1 = date.today().strftime("%d-%b-%Y")
file_date = ('File Date: ' + current_date_1)


class FooterCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []
        self.width, self.height = defaultPageSize

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            if (self._pageNumber > 1):
                self.draw_canvas(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_canvas(self, page_count):
        page = "Page %s of %s" % (self._pageNumber, page_count)
        self.saveState()
        self.setFont('Times-Roman', 10)
        self.drawString(0.9 * landscape(defaultPageSize)[0], 0.05 * landscape(defaultPageSize)[1], page)
        self.restoreState()


class Greport:

    def __init__(self, path):
        self.path = path
        self.styleSheet = getSampleStyleSheet()
        self.elements = []

        self.firstPage()
        self.report_type_1()
        # self.report_type_2()
        self.report_type_3()
        self.report_type_4()
        self.report_type_5()

        # Build
        self.doc = SimpleDocTemplate(path, pagesize=landscape(defaultPageSize),
                                     rightMargin=72,
                                     leftMargin=72,
                                     topMargin=40,
                                     bottomMargin=40)
        self.doc.multiBuild(self.elements, canvasmaker=FooterCanvas)

    def firstPage(self):
        # image
        self.elements.append(get_image('kostki.png', width=100 * mm))

        # spacer
        spacer = Spacer(0, 100)
        self.elements.append(spacer)

        # title
        fp_title = ParagraphStyle('Hed0', fontSize=20, alignment=TA_CENTER)
        fp_text = "NAZWA DEPARTEMAENTU SPOLKA ZOO"
        fp_title_summary = Paragraph(fp_text, fp_title)
        self.elements.append(fp_title_summary)

        self.elements.append(PageBreak())

    def report_type_1(self):
        # self.elements.append(PageBreak())

        elements = []

        # table title
        tbl_title = [['report_type_1', file_date]]

        table_title = Table(tbl_title,
                            colWidths=[(4 / 5 * (defaultPageSize[1] - 144)),
                                       (1 / 5 * (defaultPageSize[1] - 144))])
        table_title_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lavender),
                                        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                                        ('FONT', (0, 0), (0, 0), 'Helvetica-Bold')])

        table_title.setStyle(table_title_style)

        self.elements.append(table_title)

        # import df
        # df = pd.read_excel('Financial Sample2.xlsx', nrows=100)
        # cols = ['Segment', 'Country']  # , 'Product']
        # df_selected = df[cols]
        # lista = [df_selected.columns[:, ].values.astype(str).tolist()] + df_selected.values.tolist()
        lista = [report_A.columns[:, ].values.astype(str).tolist()] + report_A.values.tolist()

        # table style
        # tblStyle = TableStyle([('BACKGROUND', (0,0), (-1,0), colors.red),
        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)])

        # tb_wid_marg = defaultPageSize
        table = Table(lista, colWidths=[((1 / 5) * (defaultPageSize[1] - 144)),  # page size minus 2* margins
                                        ((4 / 5) * (defaultPageSize[1] - 144)),
                                        # ((1/4)*(defaultPageSize[1] - 144))
                                        ],
                      repeatRows=1)

        # table.setStyle(tblStyle)

        table_style1 = [('BACKGROUND', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)
                        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold')
                        ]
        for i, row in enumerate(range(0, table._nrows), 1):
            if i % 2 == 0:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.aliceblue))
            else:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.white))

        table.setStyle(TableStyle(table_style1))

        self.elements.append(table)

    def report_type_4(self):
        self.elements.append(PageBreak())

        elements = []

        # table title
        tbl_title = [['report_type_4', file_date]]

        table_title = Table(tbl_title,
                            colWidths=[(4 / 5 * (defaultPageSize[1] - 144)),
                                       (1 / 5 * (defaultPageSize[1] - 144))])
        table_title_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lavender),
                                        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                                        ('FONT', (0, 0), (0, 0), 'Helvetica-Bold')])

        table_title.setStyle(table_title_style)

        self.elements.append(table_title)

        # import df
        df = pd.read_excel('Financial Sample2.xlsx', nrows=100)
        cols = ['Segment', 'Country']  # , 'Product']
        df_selected = df[cols]
        lista = [df_selected.columns[:, ].values.astype(str).tolist()] + df_selected.values.tolist()

        # table style
        # tblStyle = TableStyle([('BACKGROUND', (0,0), (-1,0), colors.red),
        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)])

        # tb_wid_marg = defaultPageSize
        table = Table(lista, colWidths=[((1 / 5) * (defaultPageSize[1] - 144)),  # page size minus 2* margins
                                        ((4 / 5) * (defaultPageSize[1] - 144)),
                                        # ((1/4)*(defaultPageSize[1] - 144))
                                        ],
                      repeatRows=1)

        # table.setStyle(tblStyle)

        table_style1 = [('BACKGROUND', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)
                        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold')
                        ]
        for i, row in enumerate(range(0, table._nrows), 1):
            if i % 2 == 0:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.aliceblue))
            else:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.white))

        table.setStyle(TableStyle(table_style1))

        self.elements.append(table)

    def report_type_2(self):
        self.elements.append(PageBreak())

        elements = []

        # import df
        df = pd.read_excel('Financial Sample2.xlsx', nrows=100)
        cols = ['Segment', 'Country', 'Product']
        df_selected = df[cols]
        lista = [df_selected.columns[:, ].values.astype(str).tolist()] + df_selected.values.tolist()

        # table style
        # tblStyle = TableStyle([('BACKGROUND', (0,0), (-1,0), colors.red),
        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)])

        # tb_wid_marg = defaultPageSize
        table = Table(lista, colWidths=[((2 / 4) * (defaultPageSize[1] - 144)),  # page size minus 2* margins
                                        ((1 / 4) * (defaultPageSize[1] - 144)),
                                        ((1 / 4) * (defaultPageSize[1] - 144))],
                      repeatRows=1)

        # table.setStyle(tblStyle)

        table_style1 = [('BACKGROUND', (0, 0), (-1, 0), colors.green),
                        # ('BOX', (0, 0), (-1, -1), 0.15, colors.black),
                        # ('INNERGRID', (0, 0), (-1, -1), 0.15, colors.black)
                        ]
        for i, row in enumerate(range(0, table._nrows), 1):
            if i % 2 == 0:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.aliceblue))
            else:
                table_style1.append(('BACKGROUND', (0, i), (-1, i), colors.white))

        table.setStyle(TableStyle(table_style1))

        self.elements.append(table)

    def report_type_3(self):
        self.elements.append(PageBreak())

        elements = []

        # table title
        tbl_title = [['report_type_3', file_date]]

        table_title = Table(tbl_title,
                            colWidths=[(4 / 5 * (defaultPageSize[1] - 144)),
                                       (1 / 5 * (defaultPageSize[1] - 144))])
        table_title_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lavender),
                                        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                                        ('FONT', (0, 0), (0, 0), 'Helvetica-Bold')])

        table_title.setStyle(table_title_style)

        self.elements.append(table_title)
        # import df
        df = pd.read_excel('Financial Sample2.xlsx', nrows=100)
        cols = ['Segment', 'Country', 'Product']
        df_selected = df[cols]

        #
        for i in df_selected['Segment'].sort_values().unique():
            table1 = Table([[i]],
                           colWidths=[(defaultPageSize[1] - 144)])
            table_style1 = [('BACKGROUND', (0, 0), (-1, 0), colors.aliceblue),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold')]
            table1.setStyle(TableStyle(table_style1))
            self.elements.append(table1)
            for j in df_selected[['Product', 'Segment', 'Country']][df_selected['Segment'] == i].values.tolist():
                table2 = Table([j], colWidths=[(1 / 5 * (defaultPageSize[1] - 144)),
                                               (3 / 5 * (defaultPageSize[1] - 144)),
                                               (1 / 5 * (defaultPageSize[1] - 144))])
                self.elements.append(table2)

    def report_type_5(self):
        self.elements.append(PageBreak())

        elements = []

        # table title
        tbl_title = [['report_type_5', file_date]]

        table_title = Table(tbl_title,
                            colWidths=[(4 / 5 * (defaultPageSize[1] - 144)),
                                       (1 / 5 * (defaultPageSize[1] - 144))])
        table_title_style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lavender),
                                        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                                        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                                        ('FONT', (0, 0), (0, 0), 'Helvetica-Bold')])

        table_title.setStyle(table_title_style)

        self.elements.append(table_title)
        # import df
        df = pd.read_excel('Financial Sample2.xlsx', nrows=100)
        cols = ['Segment', 'Country', 'Product']
        df_selected = df[cols]

        #
        for i in df_selected['Segment'].sort_values().unique():
            table1 = Table([[i]],
                           colWidths=[(defaultPageSize[1] - 144)])
            table_style1 = [('BACKGROUND', (0, 0), (-1, 0), colors.aliceblue),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold')]
            table1.setStyle(TableStyle(table_style1))
            self.elements.append(table1)
            for j in df_selected[['Product', 'Segment', 'Country']][df_selected['Segment'] == i].values.tolist():
                table2 = Table([j], colWidths=[(1 / 5 * (defaultPageSize[1] - 144)),
                                               (3 / 5 * (defaultPageSize[1] - 144)),
                                               (1 / 5 * (defaultPageSize[1] - 144))])
                self.elements.append(table2)


report_pdf_save_name = os.path.join(folder_path + 'report_pdf ' + current_date_2 + '.pdf')
if __name__ == '__main__':
    report = Greport(report_pdf_save_name)

# https://stackoverflow.com/questions/17542524/pandas-dataframes-in-reportlab
# https://github.com/jurasec/python-reportlab-example/blob/master/pdf_timesheet.py#L130-L195
