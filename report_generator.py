'''
Created on Sep 27, 2018

@author: Franklin.Ventura@roche.com
'''
import time
import datetime
from itertools import zip_longest
import reportlab.pdfbase.pdfform as pdfform
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph,\
    Spacer, Flowable
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
styles = getSampleStyleSheet()
style = styles["Normal"]
style_body = styles["BodyText"]
style_body.wordWrap = 'CJK'
style_body.alignment = TA_CENTER


class SignatureDate(Flowable):
    def __init__(self, x=-10, y=10, width=280, height=30, name=1):
        Flowable.__init__(self)
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.name = name

    def draw(self):
        c = self.canv
        ### Signature ###
        c.drawString(self.x, self.y, f"{self.name} Signature")
        c.rect(self.x, self.y - 5 - self.height, self.width, self.height)
        pdfform.textFieldRelative(
            c, "textfield_Sign" + str(self.name),
            self.x, self.y - 5 - self.height, self.width, self.height,
            "")
        ### Date ###
        c.drawString(self.x + self.width + 5, self.y, "Date (DD-MMM-YYYY)")
        c.rect(self.x + self.width + 5, self.y - 5 - self.height,
               (self.width / 2.0) + 70, self.height)
        pdfform.textFieldRelative(
            c, "textfield_Date" + str(self.name),
            self.x + self.width + 5, self.y - 5 - self.height,
            (self.width / 2.0) + 70, self.height, "")
        ### Print Name ###
        c.drawString(self.x + (self.width * 1.5 + 80), self.y, "Print Name")
        c.rect(self.x + (self.width * 1.5) + 80, self.y - 5 - self.height,
               self.width, self.height)
        pdfform.textFieldRelative(
            c, "textfield_Print" + str(self.name),
            self.x + (self.width * 1.5) + 80, self.y - 5 - self.height,
            self.width, self.height, "")


class NumberedCanvas(canvas.Canvas):
    pagesize = A4

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def draw_header(self):
        width = 2 * inch
        height = 0.5 * inch
        img_x = self.pagesize[0] - width - 0.5 * inch
        img_y = self.pagesize[1] - 0.60 * inch
        self.drawImage(r"Ventana_Medical.png",
                       x=img_x, y=img_y, width=width, height=height,
                       preserveAspectRatio=True)

    def save(self):
        """add page info to each page (page x of y)"""
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            self.draw_header()
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.drawRightString((15 * self.pagesize[0] / 16.0),
                             0.1 * inch + (0.2 * inch),
                             "Page %d of %d" % (self._pageNumber, page_count))
        self.drawString((self.pagesize[0] / 16.0), 0.1 * inch + (0.2 * inch),
                        "Document Generated: " + time.strftime("%d-%b-%Y"))


class BaseReport(object):
    '''
    classdocs
    '''

    def __init__(self, doc_title, make_landscape=False):
        '''
        Constructor
        '''
        self.elements = []
        style.spaceBefore = 5
        self.doc_title = doc_title
        self.landscape = False
        if make_landscape:
            self.pagesize = landscape(A4)
            self.landscape = True
        else:
            self.pagesize = A4

    def header(self, canvas, doc):
        textobject = canvas.beginText()
        textobject.setFont(style.fontName, 16)
        textobject.setTextOrigin(
            (self.pagesize[0] / 16.0), self.pagesize[1] - (0.35 * inch))
        textobject.textLine(text=self.doc_title)
        canvas.drawText(textobject)

    def front_page_img(self, canvas, doc):  # pragma: no cover
        self.c = canvas
        width = 2.5 * inch
        height = 1 * inch
        img_x = self.pagesize[0] - width - doc.rightMargin
        img_y = self.pagesize[1] - height - doc.topMargin * 0.5 + 20
        self.c.drawImage(r"Ventana_Medical.png",
                         x=img_x, y=img_y, width=width, height=height,
                         preserveAspectRatio=True)

    def build_document(self, filepath=None):
        doc = SimpleDocTemplate(filepath,
                                title=self.doc_title,
                                pagesize=self.pagesize,
                                leftMargin=inch * 0.5,
                                rightMargin=inch * 0.5,
                                topMargin=inch * 0.5,
                                bottomMargin=inch * 0.8,)
        NumberedCanvas.pagesize = self.pagesize
        elements = self.elements.copy()
        doc.build(
            elements,
            onFirstPage=self.header,
            onLaterPages=self.header,
            canvasmaker=NumberedCanvas)

    def create_table_audit(self, audit_data, col_order=None):
        audit_table_data = []
        if audit_data:
            for audit_dict in audit_data:
                if col_order is None:
                    col_order = [value for value in audit_dict.keys()]
                if not audit_table_data:
                    audit_table_data.append(
                        [Paragraph(str(value.title()), style_body) for value in col_order])
                audit_table_data.append(
                    [Paragraph(str(audit_dict[key]), style_body) for key in col_order])
            audit_table = Table(
                audit_table_data, repeatRows=(0, ), colWidths=['15%', '15%',
                                                               '10%', '20%', '40%'])
            default_table_style = TableStyle(
                [('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                 ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                 ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                 ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
                 ('OUTLINE', (0, 0), (-1, -1), 1, colors.lightgrey),
                 ('LEFTPADDING', (0, 0), (-1, -1), 1),
                 ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                 ('TOPPADDING', (0, 0), (-1, -1), 1),
                 ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                 ])
            audit_table.setStyle(default_table_style)
            self.elements.append(Paragraph(str("Audits and Exclusions"),
                                           styles['Heading2']))
            self.elements.append(audit_table)

    def create_table_data(self, table_data, colWidths=None, col_order=None):
        first_obj = table_data[0]
        if isinstance(first_obj, dict):
            new_table_data = []
            if col_order is None:
                col_order = [value for value in first_obj.keys()]
            new_table_data.append(
                [Paragraph(str(value).replace(' ', '<br />'), style_body) for value in col_order])
            for table_dict in table_data:
                new_table_data.append(
                    [Paragraph(str(table_dict[key]), style_body) for key in col_order])
            table_data = new_table_data
        newtable = Table(table_data, repeatRows=(0, ), colWidths=colWidths)
        default_table_style = TableStyle(
                [('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                 ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                 ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                 ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
                 ('OUTLINE', (0, 0), (-1, -1), 1, colors.lightgrey),
                 ('LEFTPADDING', (0, 0), (-1, -1), 1),
                 ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                 ('TOPPADDING', (0, 0), (-1, -1), 1),
                 ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                 ])
        newtable.setStyle(default_table_style)
        self.elements.append(Paragraph(str("Dispenser Details"),
                                       styles['Heading2']))
        self.elements.append(newtable)

    def grouper(self, iterable, group_size, fillvalue=None):
        """
        Collect data into fixed-length chunks or blocks
         # grouper('ABCDEFG', 3, 'x') --> ABC DEF Gxx
        """
        args = [iter(iterable)] * group_size
        return list(zip_longest(fillvalue=fillvalue, *args))

    def create_table_summary(self, summary_dict):
        table_data = [Paragraph(f"<b>{summary}:</b> {info}", style)
                      for summary, info in summary_dict.items()]
        columns = 2
        # These help move the summary away fromthe logo
        table_data.insert(2, Paragraph("", style))
        if self.landscape:
            columns = 3
        table_data = self.grouper(table_data, columns)
        newtable = Table(table_data)
        newtable.setStyle(
            TableStyle(
                [('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                 ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                 ]))
        self.elements.append(newtable)

    def create_report(self, summary_dict, table_data, audit_data, colWidths=None):
        """
        Summary dict should be dict where key is the Name and value is the info
        """
        self.create_table_summary(summary_dict)
        self.create_table_data(table_data, colWidths)
        self.create_table_audit(audit_data)
        self.elements.append(Spacer(0, inch * 0.05))
        self.elements.append(SignatureDate(name='Operator'))
        self.elements.append(Spacer(0, inch * 0.27))
        self.elements.append(SignatureDate(name='Reviewer'))


if __name__ == "__main__":  # pragma: no cover
    doc_title = "Dispenser Functionality Tester Data Analysis"
    test = BaseReport(doc_title, make_landscape=True)
    summary_dict = {
        'Report Name': 'Ops samps_20181002_084133',
        'UserID': 'venturf2',
        'User Email': 'Franklin.Ventura@roche.com',
        'Report Created': datetime.datetime(2019, 2, 1, 8, 42, 9, 63747),
        'Analysis Version': '0.1.0',
        'Total Dispensers': 3,
        'Failing Dispensers': 2,
    }
    table_data = [
        {
            'Label': 'S2L3S3', 'Stress': 'Purple',
            'Type': '1695501 E90092',
            'Reagent': 'water',
            'Average (uL)': 91.900000000000006,
            'Std Dev (uL)': 1.01,
            'Pass/ Fail': 'Fail',
            'DFTM Complete': 6, 'DFTM Incomplete': 9,
            'Failing Reason': "12 drop(s) greater than 100",
            'Drop First Fail': 301,
        },
        {
            'Label': 'S2SL7S4', 'Stress': '25-OCT-2018TG6',
            'Type': '1695501 E90092',
            'Reagent': 'Avidin Diluent with B5 Blocker for Multimer',
            'Average (uL)': 93.6,
            'Std Dev (uL)': 8.33,
            'Pass/ Fail': 'Fail',
            'DFTM Complete': 3, 'DFTM Incomplete': 3,
            'Failing Reason': "1 drop(s) greater than 175",
            'Drop First Fail': 801,
        },
        {
            'Label': '19-OCT-2018CTR3', 'Stress': 'CONTROL',
            'Type': 'W/O OIL ON BARREL TIP',
            'Reagent': 'Avidin Diluent with B5 Blocker for Multimer',
            'Average (uL)': 91.900000000000006,
            'Std Dev (uL)': 0.97,
            'Pass/ Fail': 'Pass',
            'DFTM Complete': 2, 'DFTM Incomplete': 3,
            'Failing Reason': "N/A",
            'Drop First Fail': "N/A",
        },
    ]
    audit_data = [
        {
            'Label': '19-OCT-2018CTR3',
            'Reason': "Insufficient Prime",
            'User': 'venturf2',
            'Date': '2019-12-09 23:12:00.93 0807+00:00',
            'Description':
                "Drops excluded from analysis: 100 drops Run: 09-Nov-2018 testgroup8 100_20181109_142408",
        },
        {
            'Label': '19-OCT-2018CTR3',
            'Reason': "Update Dispenser",
            'User': 'venturf2',
            'Date': '2019-12-09 23:12:00.93 0807+00:00',
            'Description':
                "Dispenser information updated: reagent: water, stress: Purple",
        },
        {
            'Label': 'S2SL7S4',
            'Reason': "Rename Dispenser",
            'User': 'venturf2',
            'Date': '2019-12-09 23:12:00.93 0807+00:00',
            'Description':
                "Dispenser renamed from D1 to 19-OCT-2018CTR3",
        },
        ]
    # 'Label': '20%',
    # 'Stress': '20%',
    # 'Type': '20%',
    # 'Reagent': '20%',
    # 'Average (uL)': 40,
    # 'Std Dev (uL)': 40,
    # 'Pass/ Fail': 40,
    # 'DFTM Complete': 42,
    # 'DFTM Incomplete': 42,
    # 'Failing Reason': "20%",
    # 'Drop First Fail': 40,
    colWidths = ['20%', '20%', '20%', '20%', 40,
                 40, 40, 45, 45, "20%", 47]
    test.create_report(summary_dict, table_data, audit_data, colWidths)
    test.build_document(r"example.pdf")
    print("COMPLETE")
