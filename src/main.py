from docx import Document

def separador(documento):
    doc = Document(documento)
    secoes = {}
    secao_atual = None
    buffer = []


    titulos = [
        "SUMÁRIO",
        "Coleta dos dados",
        "Avaliação de Acessibilidade",
        "Referências"
    ]

    for p in doc.paragraphs:
        texto = p.text.strip()

        if not texto:
            pass

        if texto in titulos:
            if secao_atual and buffer:
                secoes[secao_atual] = "\n".join(buffer).strip()
                buffer = []
            secao_atual = texto
        else:
            buffer.append(texto)


    if secao_atual and buffer:
        secoes[secao_atual] = "\n".join(buffer).strip()

    return secoes  # <-- Faltava isso


secoes = separador("teste.docx")


for nome, conteudo in secoes.items():
    print(f"=== {nome} ===")
    print(conteudo)
    print()
