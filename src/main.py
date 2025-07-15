from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def separador_titulos(documento):
    secoes = []
    for paragrafo in documento.paragraphs:
        estilo = paragrafo.style.name
        texto = paragrafo.text.strip()
        if estilo.startswith("Heading") and texto:
            secoes.append(texto)
    return secoes

def aplicar_fonte(run):
    run.font.name = 'Arial'
    run.font.size = Pt(12)

def formatador_abnt(documento,nome):
    for section in documento.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
    for paragrafo in documento.paragraphs:
        estilo = paragrafo.style.name
        texto_original = paragrafo.text

        if estilo == "Heading 1":
            paragrafo.text = texto_original.upper()
            paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragrafo.paragraph_format.line_spacing = 1.5
            paragrafo.paragraph_format.space_after = Pt(12)
            paragrafo.paragraph_format.first_line_indent = Pt(0)
            for run in paragrafo.runs:
                run.bold = True
                aplicar_fonte(run)
                run.font.size = Pt(14)

        elif estilo == "Heading 2":
            paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragrafo.paragraph_format.line_spacing = 1.5
            paragrafo.paragraph_format.space_after = Pt(8)
            paragrafo.paragraph_format.first_line_indent = Pt(0)
            for run in paragrafo.runs:
                run.bold = True
                run.italic = True
                aplicar_fonte(run)

        elif 'REFERÊNCIAS' in texto_original.upper():
            paragrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragrafo.paragraph_format.line_spacing = 1.0
            paragrafo.paragraph_format.space_after = Pt(6)
            paragrafo.paragraph_format.left_indent = Cm(1.25)
            paragrafo.paragraph_format.first_line_indent = Cm(-1.25)
            for run in paragrafo.runs:
                run.bold = False
                run.italic = False
                aplicar_fonte(run)

        else:
            paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragrafo.paragraph_format.line_spacing = 1.5
            paragrafo.paragraph_format.space_after = Pt(0)
            paragrafo.paragraph_format.first_line_indent = Cm(1.25)
            for run in paragrafo.runs:
                run.bold = False
                run.italic = False
                aplicar_fonte(run)
        #print(paragrafo.text)
    documento.save(nome)
    print(f"\ndocumento formatado salvo como {nome}")

doc = "teste.docx"
documento = Document(doc)
sec = separador_titulos(documento)
print(f"\nSecoes: {sec}")
formatador_abnt(documento,"teste_formatado.docx")
