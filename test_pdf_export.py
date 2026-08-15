import os
import io
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def register_korean_font():
    font_path = "C:\\Windows\\Fonts\\malgun.ttf"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("Malgun", font_path))
        return "Malgun"
    return "Helvetica"

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.saveState()
        font_name = register_korean_font()
        self.setFont(font_name, 9)
        self.setFillColor(colors.HexColor("#555555"))
        
        # A4 Landscape: width = 841.89 pt, height = 595.27 pt
        # 15mm margin = 42.5 pt
        text = f"({self._pageNumber}/{page_count})"
        self.drawRightString(841.89 - 42.5, 20, text)
        self.restoreState()

def generate_pdf_report(df: pd.DataFrame) -> bytes:
    font_name = register_korean_font()
    buffer = io.BytesIO()
    
    # 15mm margin = 15 / 25.4 * 72 = 42.52 pt
    margin_pt = 42.52
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=margin_pt,
        rightMargin=margin_pt,
        topMargin=margin_pt,
        bottomMargin=margin_pt + 15
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=10
    )
    cell_style = ParagraphStyle(
        'CellText',
        fontName=font_name,
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#1F2937")
    )
    header_cell_style = ParagraphStyle(
        'HeaderCellText',
        fontName=font_name,
        fontSize=8,
        leading=10,
        textColor=colors.white,
        alignment=1 # Center
    )

    elements = []
    elements.append(Paragraph("학교생활기록부 AI 오탈자 및 기재지침 검증 리포트", title_style))
    elements.append(Spacer(1, 8))

    # Columns in df: 학번, 이름, 구분, 세부, 수정전, 수정 후, 수정해야하는 이유나 근거, 수정구분
    headers = ["학번", "이름", "구분", "세부", "수정전", "수정 후", "수정 이유/근거", "수정구분"]
    table_data = [[Paragraph(h, header_cell_style) for h in headers]]

    # A4 Landscape printable width = 841.89 - (42.52 * 2) = 756.85 pt
    col_widths = [45, 45, 40, 55, 220, 165, 140, 46.85]

    for idx, row in df.iterrows():
        r_data = [
            Paragraph(str(row.get('학번(숫자5자리)', row.get('학번', ''))), cell_style),
            Paragraph(str(row.get('이름', '')), cell_style),
            Paragraph(str(row.get('구분', '')), cell_style),
            Paragraph(str(row.get('세부', '')), cell_style),
            Paragraph(str(row.get('수정전', '')), cell_style),
            Paragraph(str(row.get('수정 후', '')), cell_style),
            Paragraph(str(row.get('수정해야하는 이유나 근거', '')), cell_style),
            Paragraph(str(row.get('수정구분', '')), cell_style),
        ]
        table_data.append(r_data)

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))

    elements.append(t)
    doc.build(elements, canvasmaker=NumberedCanvas)
    return buffer.getvalue()

# Quick test
sample_df = pd.DataFrame([
    {
        '학번(숫자5자리)': '30101',
        '이름': '고윤',
        '구분': '창체',
        '세부': '자율활동',
        '수정전': '불필요한 이메일을 영구 삭제하는 필수 과제를 제안 함.',
        '수정 후': '불필요한 이메일을 영구 삭제하는 필수 과제를 제안함.',
        '수정해야하는 이유나 근거': "'제안 함' -> '제안함' 띄어쓰기 오류 (동사 파생 접미사 -하다는 붙여 씀)",
        '수정구분': '수정 필수'
    }
])

pdf_bytes = generate_pdf_report(sample_df)
print(f"Generated PDF bytes size: {len(pdf_bytes)} bytes")
