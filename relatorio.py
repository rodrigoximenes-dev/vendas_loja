"""Gera um relatório mensal a partir de uma planilha de vendas (CSV).

Colunas obrigatórias: data (AAAA-MM-DD), produto, quantidade, valor

Uso:
    python relatorio.py                          # lê vendas.csv, usa o último mês
    python relatorio.py vendas.csv --mes 2026-10
    python relatorio.py --formato todos          # txt + Excel + PDF
    python relatorio.py --formato pdf --top 5

Excel exige openpyxl e PDF exige reportlab (pip install openpyxl reportlab).

Os arquivos Excel e PDF seguem, no que cabe a um relatório de uma página,
as normas da ABNT (NBR 14724 e NBR 6024) e as normas de apresentação tabular
do IBGE: fonte Arial 12, margens de 3 cm (superior e esquerda) e 2 cm (inferior
e direita), espaçamento 1,5, texto justificado, seções numeradas, tabelas
numeradas com título acima e fonte abaixo, paginação no canto superior direito.
"""
import argparse
import csv
import sys
from collections import defaultdict
from datetime import datetime
from math import ceil
from pathlib import Path
from xml.sax.saxutils import escape

COLUNAS = ["data", "produto", "quantidade", "valor"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
DIAS = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
        "sexta-feira", "sábado", "domingo"]
ARTIGO_DIA = ["na", "na", "na", "na", "na", "no", "no"]


class ErroPlanilha(Exception):
    """Problema na planilha de entrada, com mensagem pronta para o usuário."""


# ---------------------------------------------------------------- formatação

def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nome_mes(mes):
    """'2026-10' -> 'outubro de 2026'."""
    return f"{MESES[int(mes[5:]) - 1]} de {mes[:4]}"


def data_extenso(d):
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def plural(n, singular, plural_):
    return f"{n} {singular if n == 1 else plural_}"


def fmt_var(frac):
    """Variação com sinal: +77,2%."""
    return "n/d" if frac is None else f"{frac * 100:+.1f}%".replace(".", ",")


def mov(frac):
    """Frase de movimento: 'teve aumento de 77,2%'."""
    if frac is None:
        return "não pôde ser comparado"
    pct = f"{abs(frac) * 100:.1f}".replace(".", ",") + "%"
    if round(frac * 100, 1) == 0:
        return "permaneceu estável"
    return f"teve aumento de {pct}" if frac > 0 else f"teve redução de {pct}"


# ---------------------------------------------------------------- leitura

def num(texto):
    t = texto.strip().replace("R$", "").strip()
    if "," in t:  # formato brasileiro: 1.234,56
        t = t.replace(".", "").replace(",", ".")
    return float(t)


def _campo(linha, nome):
    return (linha.get(nome) or "").strip()


def _converter(linha):
    """Converte uma linha do CSV. Levanta ValueError com o motivo."""
    try:
        data = datetime.strptime(_campo(linha, "data"), "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"data inválida ({_campo(linha, 'data')!r}); use AAAA-MM-DD")
    produto = _campo(linha, "produto")
    if not produto:
        raise ValueError("produto em branco")
    try:
        qtd = int(num(_campo(linha, "quantidade")))
    except ValueError:
        raise ValueError(f"quantidade inválida ({_campo(linha, 'quantidade')!r})")
    try:
        valor = num(_campo(linha, "valor"))
    except ValueError:
        raise ValueError(f"valor inválido ({_campo(linha, 'valor')!r})")
    return {"data": data, "produto": produto, "qtd": qtd, "valor": valor}


def ler(caminho):
    """Lê o CSV de vendas. Levanta ErroPlanilha se o arquivo não servir."""
    arquivo = Path(caminho)
    if not arquivo.is_file():
        raise ErroPlanilha(f"Arquivo não encontrado: {caminho}")
    with open(arquivo, newline="", encoding="utf-8-sig") as f:
        amostra = f.read(2048)
        if not amostra.strip():
            raise ErroPlanilha("A planilha está vazia.")
        f.seek(0)
        sep = ";" if amostra.count(";") > amostra.count(",") else ","
        leitor = csv.DictReader(f, delimiter=sep)
        encontradas = [(n or "").strip().lower() for n in (leitor.fieldnames or [])]
        leitor.fieldnames = encontradas

        faltando = [c for c in COLUNAS if c not in encontradas]
        if faltando:
            raise ErroPlanilha(
                f"A planilha não tem {'a coluna obrigatória' if len(faltando) == 1 else 'as colunas obrigatórias'}: "
                f"{', '.join(faltando)}.\n"
                f"Colunas encontradas: {', '.join(c for c in encontradas if c) or 'nenhuma'}.\n"
                f"Colunas esperadas: {', '.join(COLUNAS)}."
            )

        vendas = []
        for i, linha in enumerate(leitor, start=2):
            if not any(str(v).strip() for v in linha.values() if v):
                continue  # linha em branco
            try:
                vendas.append(_converter(linha))
            except ValueError as e:
                print(f"Aviso: linha {i} ignorada ({e}).")
    return vendas


# ---------------------------------------------------------------- cálculo

def mes_anterior(mes):
    ano, m = int(mes[:4]), int(mes[5:])
    return f"{ano - 1}-12" if m == 1 else f"{ano}-{m - 1:02d}"


def variacao(atual, anterior):
    """Variação em fração (0.772 = +77,2%) ou None se não houver base."""
    return None if anterior == 0 else (atual - anterior) / anterior


def do_mes(vendas, mes):
    return [v for v in vendas if v["data"].strftime("%Y-%m") == mes]


def calcular(vendas, mes, top):
    """Calcula todos os indicadores do mês. Devolve None se não houver vendas."""
    linhas = do_mes(vendas, mes)
    if not linhas:
        return None

    total = sum(v["valor"] for v in linhas)
    itens = sum(v["qtd"] for v in linhas)

    por_produto = defaultdict(lambda: [0, 0.0])
    por_dia = defaultdict(float)
    for v in linhas:
        por_produto[v["produto"]][0] += v["qtd"]
        por_produto[v["produto"]][1] += v["valor"]
        por_dia[v["data"]] += v["valor"]

    produtos = sorted(
        ((nome, q, val) for nome, (q, val) in por_produto.items()),
        key=lambda p: p[1], reverse=True,
    )
    melhor_dia = max(por_dia, key=por_dia.get)

    ant = mes_anterior(mes)
    linhas_ant = do_mes(vendas, ant)
    comp = None
    if linhas_ant:
        total_ant = sum(v["valor"] for v in linhas_ant)
        itens_ant = sum(v["qtd"] for v in linhas_ant)
        comp = {
            "total": total_ant, "itens": itens_ant,
            "var_total": variacao(total, total_ant),
            "var_itens": variacao(itens, itens_ant),
        }

    return {
        "mes": mes, "total": total, "itens": itens,
        "dias": len(por_dia), "media": total / len(por_dia),
        "comp": comp, "ant": ant,
        "produtos": produtos, "ranking": produtos[:top],
        "melhor_dia": melhor_dia, "melhor_valor": por_dia[melhor_dia],
        "por_dia": sorted(por_dia.items()), "linhas": linhas,
    }


def numeracao(r):
    """Número das tabelas do relatório (a de comparação some se não houver dados)."""
    n = {"resumo": 1}
    k = 2
    if r["comp"]:
        n["comp"] = k
        k += 1
    n["prod"] = k
    return n


def narrativa(r, mes, n):
    """Parágrafos do relatório, em português correto, citando as tabelas."""
    d = r["melhor_dia"]
    wd = d.weekday()

    def brl(v):  # espaço não separável: "R$" não fica sozinho no fim da linha
        return globals()["brl"](v).replace(" ", " ")

    t = {
        "resumo": (
            f"No mês de {nome_mes(mes)}, a loja registrou faturamento total de {brl(r['total'])}, "
            f"com {plural(r['itens'], 'item vendido', 'itens vendidos')} em "
            f"{plural(r['dias'], 'dia', 'dias')} com movimento. O faturamento médio por dia com vendas "
            f"foi de {brl(r['media'])}, conforme a Tabela {n['resumo']}."
        ),
        "dia": (
            f"O maior faturamento diário ocorreu {ARTIGO_DIA[wd]} {DIAS[wd]}, "
            f"{data_extenso(d)}, com {brl(r['melhor_valor'])}."
        ),
    }
    c = r["comp"]
    if c is None:
        t["comp"] = (f"A planilha não contém vendas de {nome_mes(r['ant'])}; por isso, "
                     f"não foi possível comparar os resultados com o mês anterior.")
    else:
        t["comp"] = (
            f"Em relação a {nome_mes(r['ant'])}, o faturamento {mov(c['var_total'])} "
            f"({brl(c['total'])} para {brl(r['total'])}) e a quantidade de itens vendidos "
            f"{mov(c['var_itens'])} ({c['itens']} para {r['itens']}), conforme a Tabela {n['comp']}."
        )
    nome, q, val = r["ranking"][0]
    resto = (f"A Tabela {n['prod']} apresenta os {len(r['ranking'])} produtos de maior volume de vendas no período."
             if len(r["ranking"]) > 1 else f"A Tabela {n['prod']} apresenta esse resultado.")
    t["prod"] = (f"O produto mais vendido foi {nome}, com {plural(q, 'unidade', 'unidades')} "
                 f"e faturamento de {brl(val)}. {resto}")
    return t


# ---------------------------------------------------------------- saída: txt

def gerar_texto(r, mes):
    L = [
        f"RELATÓRIO DE VENDAS – {nome_mes(mes).capitalize()}",
        "=" * 44,
        f"Total do mês:      {brl(r['total'])}",
        f"Itens vendidos:    {r['itens']}",
        f"Dias com venda:    {r['dias']}",
        f"Média por dia:     {brl(r['media'])}",
        "",
        f"Comparação com o mês anterior ({nome_mes(r['ant'])}):",
    ]
    c = r["comp"]
    if c is None:
        L.append("  Sem dados do mês anterior na planilha.")
    else:
        L.append(f"  Total:    {brl(c['total'])} -> {brl(r['total'])}  ({fmt_var(c['var_total'])})")
        L.append(f"  Itens:    {c['itens']} -> {r['itens']}  ({fmt_var(c['var_itens'])})")
    L += ["", f"Produtos mais vendidos (top {len(r['ranking'])}):"]
    for i, (nome, q, val) in enumerate(r["ranking"], start=1):
        L.append(f"  {i}. {nome:<22} {q:>3} un.  {brl(val)}")
    d = r["melhor_dia"]
    L += ["", "Dia mais forte:",
          f"  {d.strftime('%d/%m/%Y')} ({DIAS[d.weekday()]}) - {brl(r['melhor_valor'])}"]
    return "\n".join(L)


# ---------------------------------------------------------------- saída: excel (ABNT)

def gerar_excel(r, mes, caminho, origem):
    try:
        from openpyxl import Workbook
        from openpyxl.chart import BarChart, Reference
        from openpyxl.styles import Alignment, Border, Font, Side
        from openpyxl.worksheet.page import PageMargins
        from openpyxl.worksheet.properties import PageSetupProperties
    except ImportError:
        raise SystemExit("Para gerar Excel instale: pip install openpyxl")

    A12, A12B = Font(name="Arial", size=12), Font(name="Arial", size=12, bold=True)
    A10, A10B = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    fina, media = Side(style="thin"), Side(style="medium")
    MOEDA, PCT = '"R$" #,##0.00', "+0.0%;-0.0%;0.0%"
    fonte_txt = f"Fonte: Planilha de vendas ({origem})."
    n = numeracao(r)
    textos = narrativa(r, mes, n)
    proxima = [max(n.values()) + 1]  # numeração das tabelas extras

    def configurar(ws, larguras):
        ws.sheet_view.showGridLines = False
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        # margens ABNT: 3 cm sup./esq. e 2 cm inf./dir. (1 pol. = 2,54 cm)
        ws.page_margins = PageMargins(left=3 / 2.54, right=2 / 2.54, top=3 / 2.54,
                                      bottom=2 / 2.54, header=2 / 2.54, footer=1 / 2.54)
        ws.oddHeader.right.text = "&P"  # página no canto superior direito
        ws.oddHeader.right.size, ws.oddHeader.right.font = 10, "Arial"
        for letra, w in larguras.items():
            ws.column_dimensions[letra].width = w
        return sum(larguras.values())

    def secao(ws, linha, texto):
        ws.cell(linha, 1, texto).font = A12B
        ws.row_dimensions[linha].height = 21
        return linha + 1

    def paragrafo(ws, linha, texto, ncols, largura):
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=ncols)
        c = ws.cell(linha, 1, texto)
        c.font = A12
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="justify")
        ws.row_dimensions[linha].height = ceil(len(texto) / (largura * 0.8)) * 21  # 1,5 linha
        return linha + 2

    def tabela(ws, linha, numero, titulo, cab, dados, fmts, esq=(1,)):
        """Tabela no padrão IBGE: título acima, filetes horizontais, fonte abaixo."""
        ws.cell(linha, 1, f"Tabela {numero} – {titulo}").font = A10
        linha += 1
        cab_linha = linha
        for j, nome in enumerate(cab, start=1):
            c = ws.cell(linha, j, nome)
            c.font = A10B
            c.alignment = Alignment(horizontal="left" if j in esq else "right", wrap_text=True)
            c.border = Border(top=media, bottom=fina)
        for i, row in enumerate(dados):
            linha += 1
            for j, val in enumerate(row, start=1):
                fmt = fmts[j - 1]
                if isinstance(val, tuple):
                    val, fmt = val
                c = ws.cell(linha, j, val)
                c.font = A10
                if fmt:
                    c.number_format = fmt
                c.alignment = Alignment(horizontal="left" if j in esq else "right")
                if i == len(dados) - 1:
                    c.border = Border(bottom=media)
        ws.cell(linha + 1, 1, fonte_txt).font = A10
        return linha + 3, cab_linha

    wb = Workbook()
    wb.properties.title = f"Relatório de Vendas – {nome_mes(mes)}"

    # ---- Resumo
    ws = wb.active
    ws.title = "Resumo"
    largura = configurar(ws, {"A": 32, "B": 18, "C": 18, "D": 14})
    ws.merge_cells("A1:D1")
    ws["A1"], ws["A1"].font = "RELATÓRIO DE VENDAS", A12B
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A2:D2")
    ws["A2"], ws["A2"].font = nome_mes(mes).capitalize(), A12
    ws["A2"].alignment = Alignment(horizontal="center")

    linha = secao(ws, 4, "1 RESUMO DO MÊS")
    linha = paragrafo(ws, linha, textos["resumo"], 4, largura)
    linha, _ = tabela(ws, linha, n["resumo"], f"Indicadores de {nome_mes(mes)}",
                      ["Indicador", "Valor"],
                      [["Total do mês", (r["total"], MOEDA)],
                       ["Itens vendidos", (r["itens"], "0")],
                       ["Dias com venda", (r["dias"], "0")],
                       ["Média por dia com vendas", (r["media"], MOEDA)]],
                      [None, None])

    linha = secao(ws, linha, "2 COMPARAÇÃO COM O MÊS ANTERIOR")
    linha = paragrafo(ws, linha, textos["comp"], 4, largura)
    c = r["comp"]
    if c:
        linha, cab_comp = tabela(
            ws, linha, n["comp"], f"Comparação entre {nome_mes(r['ant'])} e {nome_mes(mes)}",
            ["Indicador", nome_mes(r["ant"]).capitalize(), nome_mes(mes).capitalize(), "Variação"],
            [["Total", (c["total"], MOEDA), (r["total"], MOEDA),
              (c["var_total"] if c["var_total"] is not None else "n/d", PCT)],
             ["Itens", (c["itens"], "0"), (r["itens"], "0"),
              (c["var_itens"] if c["var_itens"] is not None else "n/d", PCT)]],
            [None] * 4)
        ws.row_dimensions[cab_comp].height = 27  # cabeçalho com nomes de mês

    linha = secao(ws, linha, f"{3 if c else 2} PRODUTOS MAIS VENDIDOS")
    linha = paragrafo(ws, linha, textos["prod"], 4, largura)
    linha, _ = tabela(ws, linha, n["prod"], f"Produtos mais vendidos em {nome_mes(mes)}",
                      ["Produto", "Quantidade", "Valor", "% do total"],
                      [[nome, q, val, val / r["total"]] for nome, q, val in r["ranking"]],
                      [None, "0", MOEDA, "0.0%"])

    linha = secao(ws, linha, f"{4 if c else 3} DIA DE MAIOR FATURAMENTO")
    paragrafo(ws, linha, textos["dia"], 4, largura)

    # ---- Produtos
    wp = wb.create_sheet("Produtos")
    configurar(wp, {"A": 30, "B": 14, "C": 16, "D": 14})
    num_t = proxima[0]
    proxima[0] += 1
    tabela(wp, 1, num_t, f"Vendas por produto em {nome_mes(mes)}",
           ["Produto", "Quantidade", "Valor", "% do total"],
           [[nome, q, val, val / r["total"]] for nome, q, val in r["produtos"]],
           [None, "0", MOEDA, "0.0%"])

    # ---- Por dia (tabela + gráfico)
    wd = wb.create_sheet("Por dia")
    configurar(wd, {"A": 16, "B": 18, "C": 16})
    num_t = proxima[0]
    proxima[0] += 1
    dados_dia = [[dia, DIAS[dia.weekday()], val] for dia, val in r["por_dia"]]
    fim, cab = tabela(wd, 1, num_t, f"Faturamento diário em {nome_mes(mes)}",
                      ["Data", "Dia da semana", "Valor"], dados_dia,
                      ["dd/mm/yyyy", None, MOEDA], esq=(1, 2))
    graf = BarChart()
    graf.type, graf.title, graf.legend = "col", "Gráfico 1 – Faturamento diário (R$)", None
    graf.x_axis.title, graf.y_axis.title = "Data", "Valor (R$)"
    graf.x_axis.delete = graf.y_axis.delete = False
    graf.x_axis.number_format = "dd/mm"
    graf.height, graf.width = 9, 15
    ultima = cab + len(dados_dia)
    graf.add_data(Reference(wd, min_col=3, min_row=cab, max_row=ultima), titles_from_data=True)
    graf.set_categories(Reference(wd, min_col=1, min_row=cab + 1, max_row=ultima))
    wd.add_chart(graf, f"A{fim}")
    wd.cell(fim + 19, 1, fonte_txt).font = A10

    # ---- Vendas do mês
    wv = wb.create_sheet("Vendas do mês")
    configurar(wv, {"A": 16, "B": 28, "C": 14, "D": 16})
    num_t = proxima[0]
    ordenadas = sorted(r["linhas"], key=lambda x: x["data"])
    _, cab = tabela(wv, 1, num_t, f"Vendas registradas em {nome_mes(mes)}",
                    ["Data", "Produto", "Quantidade", "Valor"],
                    [[v["data"], v["produto"], v["qtd"], v["valor"]] for v in ordenadas],
                    ["dd/mm/yyyy", None, "0", MOEDA], esq=(1, 2))
    wv.auto_filter.ref = f"A{cab}:D{cab + len(ordenadas)}"

    wb.save(caminho)


# ---------------------------------------------------------------- saída: pdf (ABNT)

def gerar_pdf(r, mes, caminho, origem):
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Table, TableStyle
    except ImportError:
        raise SystemExit("Para gerar PDF instale: pip install reportlab")

    FONTE, NEG = "Helvetica", "Helvetica-Bold"  # equivalentes métricos do Arial
    corpo = ParagraphStyle("corpo", fontName=FONTE, fontSize=12, leading=18,
                           alignment=TA_JUSTIFY, firstLineIndent=1.25 * cm)
    titulo = ParagraphStyle("titulo", fontName=NEG, fontSize=12, leading=18, alignment=TA_CENTER)
    subtitulo = ParagraphStyle("subtitulo", parent=titulo, fontName=FONTE, spaceAfter=18)
    secao = ParagraphStyle("secao", fontName=NEG, fontSize=12, leading=18, alignment=TA_LEFT,
                           spaceBefore=18, spaceAfter=18, keepWithNext=1)
    legenda = ParagraphStyle("legenda", fontName=FONTE, fontSize=10, leading=12,
                             spaceBefore=6, spaceAfter=4, keepWithNext=1)
    fonte_st = ParagraphStyle("fonte", fontName=FONTE, fontSize=10, leading=12,
                              spaceBefore=4, spaceAfter=6)
    fonte_txt = f"Fonte: Planilha de vendas ({escape(origem)})."

    def tabela(dados, larguras, dir_cols):
        t = Table(dados, colWidths=larguras, hAlign="LEFT")
        estilo = [
            ("FONTNAME", (0, 0), (-1, -1), FONTE), ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("LEADING", (0, 0), (-1, -1), 12), ("FONTNAME", (0, 0), (-1, 0), NEG),
            ("LINEABOVE", (0, 0), (-1, 0), 1, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
            ("LINEBELOW", (0, -1), (-1, -1), 1, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        estilo += [("ALIGN", (c, 0), (c, -1), "RIGHT") for c in dir_cols]
        t.setStyle(TableStyle(estilo))
        return t

    def bloco(numero, tit, dados, larguras, dir_cols):
        return KeepTogether([Paragraph(f"Tabela {numero} – {escape(tit)}", legenda),
                             tabela(dados, larguras, dir_cols),
                             Paragraph(fonte_txt, fonte_st)])

    n = numeracao(r)
    textos = narrativa(r, mes, n)
    c = r["comp"]
    el = [Paragraph("RELATÓRIO DE VENDAS", titulo),
          Paragraph(nome_mes(mes).capitalize(), subtitulo)]

    el.append(Paragraph("1 RESUMO DO MÊS", secao))
    el.append(Paragraph(escape(textos["resumo"]), corpo))
    el.append(bloco(n["resumo"], f"Indicadores de {nome_mes(mes)}",
                    [["Indicador", "Valor"], ["Total do mês", brl(r["total"])],
                     ["Itens vendidos", str(r["itens"])], ["Dias com venda", str(r["dias"])],
                     ["Média por dia com vendas", brl(r["media"])]],
                    [10 * cm, 6 * cm], (1,)))

    el.append(Paragraph("2 COMPARAÇÃO COM O MÊS ANTERIOR", secao))
    el.append(Paragraph(escape(textos["comp"]), corpo))
    if c:
        el.append(bloco(n["comp"], f"Comparação entre {nome_mes(r['ant'])} e {nome_mes(mes)}",
                        [["Indicador", nome_mes(r["ant"]).capitalize(), nome_mes(mes).capitalize(), "Variação"],
                         ["Total", brl(c["total"]), brl(r["total"]), fmt_var(c["var_total"])],
                         ["Itens", str(c["itens"]), str(r["itens"]), fmt_var(c["var_itens"])]],
                        [3.5 * cm, 5 * cm, 5 * cm, 2.5 * cm], (1, 2, 3)))

    el.append(Paragraph(f"{3 if c else 2} PRODUTOS MAIS VENDIDOS", secao))
    el.append(Paragraph(escape(textos["prod"]), corpo))
    dados = [["Posição", "Produto", "Quantidade", "Valor"]]
    for i, (nome, q, val) in enumerate(r["ranking"], start=1):
        dados.append([f"{i}º", nome, str(q), brl(val)])
    el.append(bloco(n["prod"], f"Produtos mais vendidos em {nome_mes(mes)}",
                    dados, [2 * cm, 7 * cm, 3 * cm, 4 * cm], (2, 3)))

    el.append(Paragraph(f"{4 if c else 3} DIA DE MAIOR FATURAMENTO", secao))
    el.append(Paragraph(escape(textos["dia"]), corpo))

    def pagina(canvas, doc):  # numeração no canto superior direito, tamanho 10
        canvas.setFont(FONTE, 10)
        canvas.drawRightString(A4[0] - 2 * cm, A4[1] - 2 * cm, str(doc.page))

    SimpleDocTemplate(
        str(caminho), pagesize=A4, title=f"Relatório de Vendas – {nome_mes(mes)}",
        leftMargin=3 * cm, topMargin=3 * cm, rightMargin=2 * cm, bottomMargin=2 * cm,
    ).build(el, onFirstPage=pagina, onLaterPages=pagina)


# ---------------------------------------------------------------- principal

def main():
    try:  # evita travar em consoles que não aceitam algum caractere
        sys.stdout.reconfigure(errors="replace")
    except AttributeError:
        pass

    p = argparse.ArgumentParser(description="Relatório mensal de vendas")
    p.add_argument("arquivo", nargs="?", default="vendas.csv")
    p.add_argument("--mes", help="AAAA-MM (padrão: último mês da planilha)")
    p.add_argument("--top", type=int, default=3, help="quantos produtos listar")
    p.add_argument("--formato", choices=["txt", "xlsx", "pdf", "todos"], default="txt",
                   help="formato de saída (padrão: txt)")
    p.add_argument("--saida", default="relatorio",
                   help="nome base dos arquivos gerados (padrão: relatorio)")
    a = p.parse_args()

    try:
        vendas = ler(a.arquivo)
    except ErroPlanilha as e:
        raise SystemExit(f"Erro: {e}")
    if not vendas:
        raise SystemExit("Erro: a planilha não tem nenhuma linha válida.")
    mes = a.mes or max(v["data"] for v in vendas).strftime("%Y-%m")

    r = calcular(vendas, mes, a.top)
    if r is None:
        raise SystemExit(f"Erro: nenhuma venda encontrada em {mes}.")

    origem = Path(a.arquivo).name
    base = Path(a.saida).with_suffix("")
    formatos = ["txt", "xlsx", "pdf"] if a.formato == "todos" else [a.formato]
    gerados = []
    for f in formatos:
        destino = base.with_suffix("." + f)
        if f == "txt":
            texto = gerar_texto(r, mes)
            print(texto)
            destino.write_text(texto + "\n", encoding="utf-8")
        elif f == "xlsx":
            gerar_excel(r, mes, destino, origem)
        else:
            gerar_pdf(r, mes, destino, origem)
        gerados.append(str(destino))
    print("\nArquivos gerados: " + ", ".join(gerados))


if __name__ == "__main__":
    main()
