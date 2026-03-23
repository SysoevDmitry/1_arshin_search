#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конвертация презентации из Markdown в PDF
Версия 3.0 - Полностью переписанная
"""

import os
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image, 
                                PageBreak, Table, TableStyle, ListFlowable, 
                                ListItem, KeepTogether, Preformatted, Flowable)
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, white, lightgrey

# Пути
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MD_FILE = os.path.join(BASE_DIR, "ПРЕЗЕНТАЦИЯ.md")
PDF_FILE = os.path.join(BASE_DIR, "ПРЕЗЕНТАЦИЯ.pdf")
IMAGES_DIR = os.path.join(BASE_DIR, "Картинки")

# Удаляем старый файл
if os.path.exists(PDF_FILE):
    os.remove(PDF_FILE)
    print(f"Удалён старый файл: {PDF_FILE}")

class ASCIIDiagram(Flowable):
    """Класс для отрисовки ASCII-диаграмм"""
    
    def __init__(self, text, font_size=6):
        Flowable.__init__(self)
        self.text = text
        self.font_size = font_size
        self.char_width = font_size * 0.6
        self.line_height = font_size * 1.15
        
    def draw(self):
        canvas = self.canv
        canvas.setFont('DejaVuMono', self.font_size)
        
        lines = self.text.split('\n')
        y = self.height - 5
        
        for line in lines:
            if line.strip():
                canvas.drawString(5, y, line)
            y -= self.line_height
    
    def wrap(self, availWidth, availHeight):
        lines = self.text.split('\n')
        max_len = max((len(line) for line in lines), default=0)
        width = max_len * self.char_width + 10
        height = len(lines) * self.line_height + 10
        return (min(width, availWidth), height)


def register_fonts():
    """Регистрация шрифтов"""
    fonts = [
        ("/usr/share/fonts/ttf/dejavu/DejaVuSans.ttf", "DejaVuSans"),
        ("/usr/share/fonts/ttf/dejavu/DejaVuSans-Bold.ttf", "DejaVuSans-Bold"),
        ("/usr/share/fonts/ttf/dejavu/DejaVuSansMono.ttf", "DejaVuMono"),
        ("/usr/share/fonts/ttf/dejavu/DejaVuSansMono-Bold.ttf", "DejaVuMono-Bold"),
    ]
    
    for path, name in fonts:
        if os.path.exists(path):
            pdfmetrics.registerFont(TTFont(name, path))
            print(f"✓ {name}")


def create_styles():
    """Создание стилей"""
    styles = getSampleStyleSheet()
    
    return {
        'Normal': ParagraphStyle(
            'Normal',
            parent=styles['Normal'],
            fontName='DejaVuSans',
            fontSize=10,
            leading=13,
            spaceAfter=6,
            spaceBefore=3,
        ),
        'H1': ParagraphStyle(
            'H1',
            parent=styles['Heading1'],
            fontName='DejaVuSans-Bold',
            fontSize=16,
            leading=20,
            textColor=HexColor('#1a3a52'),
            spaceAfter=14,
            spaceBefore=18,
            alignment=TA_CENTER,
        ),
        'H2': ParagraphStyle(
            'H2',
            parent=styles['Heading2'],
            fontName='DejaVuSans-Bold',
            fontSize=13,
            leading=16,
            textColor=HexColor('#2c5282'),
            spaceAfter=10,
            spaceBefore=14,
        ),
        'H3': ParagraphStyle(
            'H3',
            parent=styles['Heading3'],
            fontName='DejaVuSans-Bold',
            fontSize=11,
            leading=14,
            textColor=HexColor('#2b6cb0'),
            spaceAfter=6,
            spaceBefore=10,
        ),
        'Code': ParagraphStyle(
            'Code',
            fontName='DejaVuMono',
            fontSize=8,
            leading=11,
            textColor=HexColor('#c53030'),
            backColor=HexColor('#f7fafc'),
            borderPadding=(8, 8, 8, 8),
            spaceAfter=10,
            spaceBefore=6,
        ),
        'Caption': ParagraphStyle(
            'Caption',
            fontName='DejaVuSans',
            fontSize=8,
            leading=10,
            textColor=HexColor('#718096'),
            alignment=TA_CENTER,
            spaceAfter=12,
        ),
    }


def clean_text(text):
    """Очистка от эмодзи и Markdown"""
    emoji_pattern = re.compile("["
        u"\U0001F600-\U0001F64F"
        u"\U0001F300-\U0001F5FF"
        u"\U0001F680-\U0001F6FF"
        u"\U0001F1E0-\U0001F1FF"
        u"\U00002702-\U000027B0"
        u"\U0001F900-\U0001F9FF"
        "]+", flags=re.UNICODE)
    text = emoji_pattern.sub('', text)
    text = text.replace('**', '').replace('*', '').replace('`', '')
    return text.strip()


def parse_markdown(filepath):
    """Парсинг Markdown"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    sections = []
    lines = content.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        if not line.strip() or line.startswith('---'):
            i += 1
            continue
        
        # H1
        if line.startswith('# ') and not line.startswith('##'):
            sections.append(('h1', line[2:].strip()))
        # H2
        elif line.startswith('## '):
            sections.append(('h2', line[3:].strip()))
        # H3
        elif line.startswith('### '):
            sections.append(('h3', line[4:].strip()))
        # Изображение
        elif line.startswith('!['):
            match = re.match(r'!\[([^\]]*)\]\(([^)]+)\)', line)
            if match:
                sections.append(('image', match.groups()))
        # ASCII-диаграмма
        elif any(c in line for c in '─│┌┐└┘├┬┴╔╗╚╝╠╣╩╬'):
            ascii_lines = []
            start_i = i
            while i < len(lines) and any(c in lines[i] for c in '─│┌┐└├┤┬┴╔╗╝╠╩╬'):
                ascii_lines.append(lines[i])
                i += 1
            if ascii_lines:
                sections.append(('ascii', '\n'.join(ascii_lines), start_i))
            i -= 1
        # Таблица
        elif line.startswith('|') and '---' not in line:
            table_rows = []
            start_i = i
            while i < len(lines) and lines[i].startswith('|') and '---' not in lines[i]:
                row = lines[i].strip('|').split('|')
                row = [clean_text(c) for c in row]
                table_rows.append(row)
                i += 1
            if len(table_rows) > 1:
                sections.append(('table', table_rows, start_i))
            i -= 1
        # Список
        elif line.startswith('- ') or line.startswith('* '):
            list_items = []
            while i < len(lines) and (lines[i].startswith('- ') or lines[i].startswith('* ')):
                list_items.append(clean_text(lines[i][2:]))
                i += 1
            sections.append(('list', list_items))
            i -= 1
        # Код
        elif line.startswith('```'):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].startswith('```'):
                code_lines.append(lines[i])
                i += 1
            sections.append(('code', '\n'.join(code_lines)))
        # Текст
        elif line.strip():
            sections.append(('text', clean_text(line)))
        
        i += 1
    
    return sections


def create_table(data):
    """Создание таблицы"""
    if not data or len(data) < 2:
        return None
    
    num_cols = len(data[0])
    num_rows = len(data)
    
    # Ширина страницы A4 минус поля = ~17cm
    page_width = A4[0] - 3*cm
    col_width = page_width / num_cols
    
    # Ограничиваем минимальную и максимальную ширину
    col_width = max(1.5*cm, min(col_width, 5*cm))
    colWidths = [col_width] * num_cols
    
    table = Table(data, colWidths=colWidths, repeatRows=1)
    
    table.setStyle(TableStyle([
        # Заголовок
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1a3a52')),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, 0), 'DejaVuSans-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Тело
        ('FONTNAME', (0, 1), (-1, -1), 'DejaVuSans'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        
        # Сетка
        ('GRID', (0, 0), (-1, -1), 0.3, HexColor('#cbd5e0')),
        ('LINEBELOW', (0, 0), (-1, 0), 1.5, HexColor('#1a3a52')),
        
        # Чередование
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [white, HexColor('#f7fafc')]),
    ]))
    
    return table


def create_pdf():
    """Создание PDF"""
    print("\n=== Создание PDF ===\n")
    
    register_fonts()
    styles = create_styles()
    sections = parse_markdown(MD_FILE)
    
    # Документ
    doc = SimpleDocTemplate(
        PDF_FILE,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm,
    )
    
    story = []
    first = True
    page_items = 0
    items_since_break = 0
    
    for idx, section in enumerate(sections):
        section_type = section[0]
        content = section[1]
        
        # Добавляем разрыв страницы перед заголовками H1/H2 если это не первый элемент
        if section_type in ('h1', 'h2') and not first and items_since_break > 3:
            story.append(PageBreak())
            items_since_break = 0
        
        if section_type == 'h1':
            if first:
                # Титульник
                story.append(Spacer(1, 4*cm))
                story.append(Paragraph("ФГИС АРШИН", styles['H1']))
                story.append(Spacer(1, 1*cm))
                story.append(Paragraph("Презентация приложения", styles['H2']))
                story.append(Spacer(1, 2*cm))
                story.append(Paragraph("Версия 2.12", styles['Normal']))
                story.append(Paragraph("Март 2026", styles['Normal']))
                story.append(PageBreak())
                first = False
                items_since_break = 0
            else:
                story.append(Paragraph(content, styles['H1']))
                items_since_break += 1
        
        elif section_type == 'h2':
            story.append(Spacer(1, 10))
            story.append(Paragraph(content, styles['H2']))
            items_since_break += 1
        
        elif section_type == 'h3':
            story.append(Spacer(1, 6))
            story.append(Paragraph(content, styles['H3']))
            items_since_break += 1
        
        elif section_type == 'text':
            story.append(Paragraph(content, styles['Normal']))
            items_since_break += 1
        
        elif section_type == 'list':
            items = [ListItem(Paragraph(t, styles['Normal']), leftIndent=10) for t in content]
            story.append(ListFlowable(items, bulletType='bullet'))
            story.append(Spacer(1, 6))
            items_since_break += 1
        
        elif section_type == 'code':
            code = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Preformatted(code, styles['Code']))
            items_since_break += 1
        
        elif section_type == 'ascii':
            story.append(Spacer(1, 6))
            story.append(KeepTogether(ASCIIDiagram(content, font_size=5.5)))
            story.append(Spacer(1, 6))
            items_since_break += 1
        
        elif section_type == 'table':
            table = create_table(content)
            if table:
                story.append(Spacer(1, 8))
                story.append(KeepTogether(table))
                story.append(Spacer(1, 8))
                items_since_break += 1
        
        elif section_type == 'image':
            alt, src = content
            img_path = os.path.join(BASE_DIR, src)
            
            if os.path.exists(img_path):
                try:
                    # Подгонка под ширину страницы
                    max_w = A4[0] - 3*cm
                    max_h = 10*cm
                    
                    img = Image(img_path, width=max_w, height=max_h)
                    img.hAlign = 'CENTER'
                    
                    story.append(Spacer(1, 8))
                    story.append(img)
                    story.append(Paragraph(alt, styles['Caption']))
                    items_since_break += 1
                except Exception as e:
                    print(f"⚠ {img_path}: {e}")
    
    doc.build(story)
    print(f"\n✓ PDF создан: {PDF_FILE}")


if __name__ == '__main__':
    create_pdf()
