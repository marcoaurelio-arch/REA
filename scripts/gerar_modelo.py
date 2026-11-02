"""
Gerador do Modelo Padronizado de Viabilidade de Strip Mall (Nexa Malls).

Gera modelo/Viabilidade_Strip_Mall.xlsx com fórmulas vivas:
    Instruções | Resumo | Premissas | Mix | Fluxo | Sensibilidade

Uso:
    python scripts/gerar_modelo.py [caminho_saida.xlsx]

Após gerar, recalcule (LibreOffice) para gravar valores em cache, ou apenas
abra no Excel (o arquivo força recálculo completo na abertura).
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- estilos
FONT = "Arial"
AZUL = "0000FF"      # input
PRETO = "000000"     # fórmula
VERDE = "008000"     # link para outra aba
NAVY = "1F3864"
CINZA = "F2F2F2"
AMARELO = "FFFF00"

F_TITULO = Font(name=FONT, size=14, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=9, italic=True, color="595959")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SEC = Font(name=FONT, size=10, bold=True, color=NAVY)
F_IN = Font(name=FONT, size=10, color=AZUL)
F_FX = Font(name=FONT, size=10, color=PRETO)
F_FXB = Font(name=FONT, size=10, color=PRETO, bold=True)
F_LK = Font(name=FONT, size=10, color=VERDE)
F_LKB = Font(name=FONT, size=10, color=VERDE, bold=True)
F_TXT = Font(name=FONT, size=10)
F_NOTE = Font(name=FONT, size=9, color="595959")

FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_KEY = PatternFill("solid", fgColor=AMARELO)
FILL_TOT = PatternFill("solid", fgColor=CINZA)

THIN = Side(style="thin", color="BFBFBF")
B_TOP = Border(top=Side(style="thin", color="000000"))

FMT_RS = 'R$ #,##0;(R$ #,##0);"-"'
FMT_RS2 = 'R$ #,##0.00;(R$ #,##0.00);"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_PCT2 = '0.00%;(0.00%);"-"'
FMT_NUM = '#,##0;(#,##0);"-"'
FMT_DEC = '0.00;(0.00);"-"'
FMT_X = '0.00"x";(0.00"x");"-"'
FMT_ANO = "0"
FMT_ANOS = '0.0" anos";(0.0" anos");"-"'
FMT_IDX = "0.0000"

ANOS = list(range(0, 13))           # Ano 0..12 (12 só para NOI forward)
COL0 = 3                            # coluna C = Ano 0
YCOLS = [get_column_letter(COL0 + i) for i in ANOS]   # C..O
FIRST, LAST = YCOLS[0], YCOLS[-1]
CEN = [("B", "Pessimista"), ("C", "Realista"), ("D", "Otimista")]


def cell(ws, ref, value, font=F_FX, fmt=None, fill=None, bold=None, align=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


def section(ws, row, text, ncols):
    for i in range(1, ncols + 1):
        ws.cell(row=row, column=i).fill = FILL_SEC
    cell(ws, f"A{row}", text, F_SEC, fill=FILL_SEC)


def header(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ===================================================== PREMISSAS (inputs)
# (chave, rótulo, unidade, formato, valores, nota)
# valores: tupla (pess, real, otim) => varia por cenário (azul nos 3)
#          valor único => input só em Realista; Pess/Otim = fórmula "=C"
PREMISSAS = [
    ("SEC", "1. IDENTIFICAÇÃO DO PROJETO"),
    ("nome", "Nome do projeto", "texto", None, "Strip Mall Exemplo",
     "Preencher. Exemplo ilustrativo."),
    ("local", "Localização (cidade/UF, endereço)", "texto", None, "Cidade/UF - preencher",
     "Preencher."),
    ("ano_base", "Ano-base (Ano 0 = compra do terreno)", "ano", FMT_ANO, 2027,
     "Valores monetários de entrada estão a preços do Ano 0."),

    ("SEC", "2. TERRENO E PRODUTO"),
    ("terreno_m2", "Área do terreno", "m²", FMT_NUM, 7000,
     "Fato a confirmar via matrícula. Exemplo ilustrativo."),
    ("ca_max", "Coeficiente de aproveitamento máximo (CA)", "x", FMT_DEC, 1.0,
     "Confirmar na lei de zoneamento / certidão de uso do solo."),
    ("to_max", "Taxa de ocupação máxima (TO)", "%", FMT_PCT, 0.60,
     "Confirmar na lei de zoneamento."),
    ("pavimentos", "Nº de pavimentos de lojas", "un", FMT_NUM, 1,
     "Strip mall típico: térreo (1) ou térreo + mezanino (2)."),
    ("eficiencia", "Eficiência ABL / Área construída", "%", FMT_PCT, 0.88,
     "Premissa. Áreas comuns, sanitários, técnica. Típico 85%-92%."),
    ("vagas", "Vagas de estacionamento planejadas", "un", FMT_NUM, 115,
     "Premissa. Ver indicador vagas/100 m² ABL na aba Mix."),
    ("area_vaga", "Área por vaga (inclui circulação)", "m²/vaga", FMT_NUM, 25,
     "Premissa usual 22-28 m²/vaga a céu aberto."),

    ("SEC", "3. CUSTOS DE IMPLANTAÇÃO (preços do Ano 0)"),
    ("terreno_rs", "Custo do terreno", "R$", FMT_RS, 3500000,
     "Exemplo: R$ 500/m². Em permuta, informar o valor econômico equivalente."),
    ("itbi", "ITBI + escritura + registro", "% terreno", FMT_PCT, 0.03,
     "Premissa. ITBI varia por município (ex.: 2%-3%) + cartório."),
    ("obra_m2", "Custo de obra (área construída)", "R$/m²", FMT_RS, (4500, 3800, 3500),
     "Referência Nexa: R$ 3.500-5.000/m² conforme padrão. Validar com orçamento."),
    ("infra_m2", "Infraestrutura externa (pavimento, drenagem, paisagismo)", "R$/m² ext.",
     FMT_RS, 250, "Premissa. Aplicada sobre área externa (terreno - projeção)."),
    ("projetos", "Projetos, aprovações e consultorias", "% obra", FMT_PCT, 0.05,
     "Premissa. Típico 4%-7% do custo de obra."),
    ("taxas", "Taxas, licenças, outorga e contrapartidas", "R$", FMT_RS, 150000,
     "Premissa. Levantar junto à prefeitura (outorga onerosa, EIV, alvarás)."),
    ("gerenc", "Gerenciamento / fiscalização de obra", "% obra", FMT_PCT, 0.03,
     "Premissa."),
    ("conting", "Contingência", "% obra+infra", FMT_PCT, (0.10, 0.07, 0.05),
     "Premissa. Maior quanto menor a maturidade do orçamento."),
    ("ti_m2", "Contribuição a lojistas (TI / allowance)", "R$/m² ABL", FMT_RS, 0,
     "Premissa. Usar quando houver entrega acima do 'casco'."),
    ("comissao", "Comissão de comercialização", "nº aluguéis", FMT_DEC, 1.0,
     "Premissa. Sobre aluguel mensal da ABL ocupada estabilizada."),
    ("marketing", "Marketing de lançamento / pré-comercialização", "R$", FMT_RS, 100000,
     "Premissa."),
    ("p_obra_a0", "% da obra desembolsada no Ano 0", "%", FMT_PCT, 0.20,
     "Restante no Ano 1. Operação inicia no Ano 2."),
    ("incc", "Reajuste de custos no Ano 1 (INCC)", "% a.a.", FMT_PCT, 0.05,
     "Premissa. Aplicado ao desembolso do Ano 1."),

    ("SEC", "4. RECEITAS"),
    ("fator_alug", "Fator sobre o aluguel do Mix", "x", FMT_DEC, (0.90, 1.00, 1.08),
     "Aluguel/m² base vem da aba Mix. Pessimista = -10%, Otimista = +8%."),
    ("reajuste", "Reajuste anual dos aluguéis e custos operacionais", "% a.a.", FMT_PCT,
     0.045, "Premissa (IPCA/IGP-M). Fluxo projetado em termos nominais."),
    ("occ1", "Ocupação no 1º ano de operação (Ano 2)", "% ABL", FMT_PCT, (0.60, 0.75, 0.85),
     "Premissa. Depende do nível de pré-locação na entrega."),
    ("occ2", "Ocupação no 2º ano de operação (Ano 3)", "% ABL", FMT_PCT, (0.80, 0.88, 0.93),
     "Premissa."),
    ("vacancia", "Vacância estrutural estabilizada (a partir do Ano 4)", "% ABL", FMT_PCT,
     (0.15, 0.10, 0.05), "Referência Nexa conservadora: 10%-20%."),
    ("carencia", "Carência média no 1º ano de operação", "meses", FMT_DEC, (4, 3, 2),
     "Premissa. Aplicada sobre a ABL ocupada no Ano 2."),
    ("outras", "Outras receitas (mídia, quiosques, estacionamento)", "% aluguel", FMT_PCT,
     (0.01, 0.02, 0.03), "Premissa."),
    ("inad", "Inadimplência", "% receita", FMT_PCT, (0.05, 0.03, 0.02), "Premissa."),

    ("SEC", "5. DESPESAS DO PROPRIETÁRIO (não repassadas ao lojista)"),
    ("adm", "Taxa de administração / gestão", "% receita", FMT_PCT, 0.10,
     "Referência Nexa: 10%-15%."),
    ("nrep", "Custos operacionais não repassáveis", "% receita", FMT_PCT, (0.06, 0.05, 0.04),
     "Referência Nexa: 3%-7% (seguro, jurídico, auditoria, manutenção estrutural)."),
    ("custo_vago", "Condomínio + IPTU de lojas vagas", "R$/m²/mês", FMT_RS2, 18,
     "Premissa. Proprietário arca com encargos da ABL vaga."),
    ("reserva", "Reserva técnica (CAPEX de manutenção)", "% receita", FMT_PCT, 0.01,
     "Referência Nexa: 0,5%-2%."),
    ("tributos", "Tributos sobre a receita de locação", "% receita", FMT_PCT2, 0.1133,
     "Premissa: PJ lucro presumido (PIS/COFINS 3,65% + IRPJ/CSLL s/ presunção 32% = "
     "7,68%; sem adicional de IR). Validar com contador; reforma tributária (IBS/CBS) "
     "pode alterar a partir de 2027. Use 0% para análise pré-impostos."),

    ("SEC", "6. SAÍDA E CRITÉRIOS DE RETORNO"),
    ("ano_saida", "Ano de saída (venda do ativo)", "ano (4 a 11)", FMT_ANO, 11,
     "11 = 10 anos de operação. Valor de saída = NOI do ano seguinte / cap rate."),
    ("cap_saida", "Cap rate de saída", "% a.a.", FMT_PCT2, (0.105, 0.095, 0.09),
     "HIPÓTESE - validar com transações comparáveis e taxa de juros vigente."),
    ("custo_venda", "Custos de venda (corretagem, assessoria)", "% valor", FMT_PCT, 0.03,
     "Premissa."),
    ("ir_ganho", "IR/CSLL sobre ganho de capital na venda", "% ganho", FMT_PCT, 0.34,
     "Premissa: PJ (IRPJ 25% + CSLL 9%). Estruturas via FII/PF têm regras diferentes."),
    ("tma", "Taxa mínima de atratividade (desconto, nominal)", "% a.a.", FMT_PCT, 0.14,
     "Premissa: custo de oportunidade (CDI + prêmio de risco). Ajustar ao momento."),
    ("spread", "Spread mínimo: Yield on Cost - cap rate de saída", "p.p.", FMT_PCT2, 0.015,
     "Opinião estratégica: prêmio de desenvolvimento usual de 1,0-2,0 p.p."),
]

P = {}  # chave -> linha na aba Premissas


def build_premissas(wb):
    ws = wb.create_sheet("Premissas")
    cell(ws, "A1", "PREMISSAS DO ESTUDO", F_TITULO)
    cell(ws, "A2", "Azul = input editável | Preto = fórmula (padrão = Realista; "
         "sobrescreva se quiser) | Amarelo = premissa-chave por cenário", F_SUB)
    header(ws, 4, ["Premissa", "Pessimista", "Realista", "Otimista", "Unidade",
                   "Nota / fonte / classificação"])
    r = 5
    for item in PREMISSAS:
        if item[0] == "SEC":
            section(ws, r, item[1], 6)
            r += 1
            continue
        key, label, unit, fmt, vals, note = item
        P[key] = r
        cell(ws, f"A{r}", label, F_TXT)
        cell(ws, f"E{r}", unit, F_NOTE)
        cell(ws, f"F{r}", note, F_NOTE)
        if isinstance(vals, tuple):
            for (col, _), v in zip(CEN, vals):
                cell(ws, f"{col}{r}", v, F_IN, fmt, fill=FILL_KEY)
        elif isinstance(vals, str):
            cell(ws, f"C{r}", vals, F_IN)
        else:
            cell(ws, f"C{r}", vals, F_IN, fmt)
            cell(ws, f"B{r}", f"=C{r}", F_FX, fmt)
            cell(ws, f"D{r}", f"=C{r}", F_FX, fmt)
        r += 1

    dv = DataValidation(type="whole", operator="between", formula1="4", formula2="11",
                        showErrorMessage=True, error="Ano de saída deve estar entre 4 e 11.")
    ws.add_data_validation(dv)
    for col in "BCD":
        dv.add(f"{col}{P['ano_saida']}")

    ws.column_dimensions["A"].width = 58
    for col in "BCD":
        ws.column_dimensions[col].width = 15
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 95
    ws.freeze_panes = "B5"
    return ws


def pr(key, col):
    """Referência absoluta a uma premissa para o cenário (coluna B/C/D)."""
    return f"Premissas!${col}${P[key]}"


# ============================================================ MIX + ÁREAS
MIX_ROWS = [
    ("Âncora - farmácia / drogaria", "Rede nacional ou regional", 1, 600, 55, 10),
    ("Âncora - academia", "Academia de bairro / low cost", 1, 800, 35, 10),
    ("Alimentação", "Restaurantes, hamburgueria, açaí, café", 4, 120, 85, 5),
    ("Conveniência", "Padaria, hortifruti, empório", 2, 150, 70, 5),
    ("Serviços", "Lavanderia, pet shop, barbearia, chaveiro", 4, 60, 90, 5),
    ("Saúde e bem-estar", "Clínicas, odontologia, laboratório", 3, 80, 80, 5),
    ("Varejo de vizinhança", "Ótica, papelaria, presentes", 3, 70, 75, 5),
]
MIX_FIRST, MIX_LAST = 5, 16          # 7 exemplos + 5 linhas livres
MIX_TOT = 17
M = {}  # chave -> referência na aba Mix


def build_mix(wb):
    ws = wb.create_sheet("Mix")
    cell(ws, "A1", "MIX COMERCIAL E QUADRO DE ÁREAS", F_TITULO)
    cell(ws, "A2", "Edite segmentos, nº de lojas, área e aluguel (azul). Linhas 12-16 "
         "livres para novos segmentos. Linhas 5-11 = exemplo ilustrativo.", F_SUB)
    header(ws, 4, ["Segmento", "Tipologia / exemplos", "Nº de lojas", "Área média (m²)",
                   "ABL (m²)", "Aluguel (R$/m²/mês)", "Receita mensal (R$)", "% ABL",
                   "% Receita", "Prazo típico (anos)"])
    for i, r in enumerate(range(MIX_FIRST, MIX_LAST + 1)):
        if i < len(MIX_ROWS):
            seg, tip, n, a, al, prazo = MIX_ROWS[i]
        else:
            seg, tip, n, a, al, prazo = None, None, None, None, None, None
        cell(ws, f"A{r}", seg, F_IN)
        cell(ws, f"B{r}", tip, F_IN)
        cell(ws, f"C{r}", n, F_IN, FMT_NUM)
        cell(ws, f"D{r}", a, F_IN, FMT_NUM)
        cell(ws, f"E{r}", f"=C{r}*D{r}", F_FX, FMT_NUM)
        cell(ws, f"F{r}", al, F_IN, FMT_RS2)
        cell(ws, f"G{r}", f"=E{r}*F{r}", F_FX, FMT_RS)
        cell(ws, f"H{r}", f"=IFERROR(E{r}/$E${MIX_TOT},0)", F_FX, FMT_PCT)
        cell(ws, f"I{r}", f"=IFERROR(G{r}/$G${MIX_TOT},0)", F_FX, FMT_PCT)
        cell(ws, f"J{r}", prazo, F_IN, FMT_NUM)
    t = MIX_TOT
    rng = lambda c: f"{c}{MIX_FIRST}:{c}{MIX_LAST}"
    for c in "ABCDEFGHIJ":
        ws[f"{c}{t}"].fill = FILL_TOT
        ws[f"{c}{t}"].border = B_TOP
    cell(ws, f"A{t}", "TOTAL / MÉDIA PONDERADA", F_FXB, fill=FILL_TOT)
    cell(ws, f"C{t}", f"=SUM({rng('C')})", F_FXB, FMT_NUM, fill=FILL_TOT)
    cell(ws, f"E{t}", f"=SUM({rng('E')})", F_FXB, FMT_NUM, fill=FILL_TOT)
    cell(ws, f"F{t}", f"=IFERROR(G{t}/E{t},0)", F_FXB, FMT_RS2, fill=FILL_TOT)
    cell(ws, f"G{t}", f"=SUM({rng('G')})", F_FXB, FMT_RS, fill=FILL_TOT)
    cell(ws, f"H{t}", f"=SUM({rng('H')})", F_FXB, FMT_PCT, fill=FILL_TOT)
    cell(ws, f"I{t}", f"=SUM({rng('I')})", F_FXB, FMT_PCT, fill=FILL_TOT)
    M["abl"] = f"Mix!$E${t}"
    M["aluguel"] = f"Mix!$F${t}"
    M["lojas"] = f"Mix!$C${t}"
    M["rec_mensal"] = f"Mix!$G${t}"

    # % de ABL e receita em âncoras (linhas cujo segmento começa com "Âncora")
    r = t + 1
    cell(ws, f"A{r}", "Participação das âncoras (segmentos iniciados por 'Âncora')", F_TXT)
    cell(ws, f"H{r}", f'=SUMIF({rng("A")},"Âncora*",{rng("H")})', F_FX, FMT_PCT)
    cell(ws, f"I{r}", f'=SUMIF({rng("A")},"Âncora*",{rng("I")})', F_FX, FMT_PCT)
    M["ancora_abl"] = f"Mix!$H${r}"
    M["ancora_rec"] = f"Mix!$I${r}"

    # ---- Quadro de áreas
    r = t + 3
    section(ws, r, "QUADRO DE ÁREAS (calculado)", 10)
    header(ws, r + 1, ["Item", "", "Valor", "Unidade"])
    pc = "C"  # parâmetros físicos não variam por cenário: usa Realista
    rows = [
        ("abl", "ABL total", f"={M['abl'].replace('Mix!', '')}", "m²", FMT_NUM),
        ("ac", "Área construída (ABL / eficiência)",
         f"=IFERROR({{abl}}/{pr('eficiencia', pc)},0)", "m²", FMT_NUM),
        ("proj", "Projeção da edificação (AC / pavimentos)",
         f"=IFERROR({{ac}}/{pr('pavimentos', pc)},0)", "m²", FMT_NUM),
        ("ext", "Área externa (terreno - projeção)",
         f"=MAX(0,{pr('terreno_m2', pc)}-{{proj}})", "m²", FMT_NUM),
        ("pot_constr", "Potencial construtivo (terreno x CA)",
         f"={pr('terreno_m2', pc)}*{pr('ca_max', pc)}", "m²", FMT_NUM),
        ("proj_max", "Projeção máxima (terreno x TO)",
         f"={pr('terreno_m2', pc)}*{pr('to_max', pc)}", "m²", FMT_NUM),
        ("area_vagas", "Área necessária para vagas",
         f"={pr('vagas', pc)}*{pr('area_vaga', pc)}", "m²", FMT_NUM),
        ("vagas100", "Vagas por 100 m² de ABL",
         f"=IFERROR({pr('vagas', pc)}/{{abl}}*100,0)", "vagas", FMT_DEC),
        ("ca_uso", "CA utilizado (AC / terreno)",
         f"=IFERROR({{ac}}/{pr('terreno_m2', pc)},0)", "x", FMT_DEC),
        ("to_uso", "TO utilizada (projeção / terreno)",
         f"=IFERROR({{proj}}/{pr('terreno_m2', pc)},0)", "%", FMT_PCT),
    ]
    rr = r + 2
    local = {}
    for key, lab, f, unit, fmt in rows:
        local[key] = f"C{rr}"
        M[key] = f"Mix!$C${rr}"
        rr += 1
    rr = r + 2
    for key, lab, f, unit, fmt in rows:
        f = f.format(**local)
        font = F_FX
        cell(ws, f"A{rr}", lab, F_TXT)
        cell(ws, f"C{rr}", f, font, fmt)
        cell(ws, f"D{rr}", unit, F_NOTE)
        rr += 1

    # ---- Checagens
    r = rr + 1
    section(ws, r, "CHECAGENS DE PRODUTO", 10)
    header(ws, r + 1, ["Checagem", "", "Status", "Critério"])
    checks = [
        ("Área construída dentro do potencial construtivo",
         f'=IF({local["ac"]}<={local["pot_constr"]},"OK","ALERTA")', "AC <= terreno x CA"),
        ("Projeção dentro da taxa de ocupação",
         f'=IF({local["proj"]}<={local["proj_max"]},"OK","ALERTA")', "Projeção <= terreno x TO"),
        ("Área externa comporta as vagas",
         f'=IF({local["ext"]}>={local["area_vagas"]},"OK","ALERTA")',
         "Área externa >= vagas x área/vaga"),
        ("Vagas por 100 m² de ABL >= 3,0",
         f'=IF({local["vagas100"]}>=3,"OK","ALERTA")',
         "Opinião estratégica: < 3 vagas/100 m² compromete conveniência"),
        ("Existe âncora no mix (>= 25% da ABL)",
         f'=IF({M["ancora_abl"].replace("Mix!", "")}>=0.25,"OK","ALERTA")',
         "Opinião estratégica: âncora gera fluxo recorrente"),
        ("Nenhum segmento acima de 40% da receita",
         f'=IF(MAX({rng("I")})<=0.4,"OK","ALERTA")', "Risco de concentração em operador"),
    ]
    rr = r + 2
    M["checks_first"] = rr
    for lab, f, crit in checks:
        cell(ws, f"A{rr}", lab, F_TXT)
        cell(ws, f"C{rr}", f, F_FXB, align=Alignment(horizontal="center"))
        cell(ws, f"D{rr}", crit, F_NOTE)
        rr += 1
    M["checks_last"] = rr - 1
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    bad_fill = PatternFill("solid", fgColor="FFC7CE")
    rng_chk = f"C{M['checks_first']}:C{M['checks_last']}"
    ws.conditional_formatting.add(rng_chk, CellIsRule(operator="equal", formula=['"OK"'],
                                                      fill=ok_fill))
    ws.conditional_formatting.add(rng_chk, CellIsRule(operator="equal",
                                                      formula=['"ALERTA"'], fill=bad_fill))
    M["checks_rows"] = [(lab, f"Mix!C{M['checks_first'] + i}")
                        for i, (lab, _, _) in enumerate(checks)]

    widths = [40, 40, 12, 14, 12, 16, 18, 10, 10, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    return ws


# ================================================================== FLUXO
BLOCK = 60
F = {}  # (cenário col, chave) -> linha


def build_fluxo(wb):
    ws = wb.create_sheet("Fluxo")
    cell(ws, "A1", "FLUXO DE CAIXA ANUAL DO PROJETO (nominal, desalavancado)", F_TITULO)
    cell(ws, "A2", "Ano 0 = aquisição e início das obras | Ano 1 = obras | Ano 2 = início "
         "da operação. Coluna do Ano 12 serve apenas para o NOI forward da saída.", F_SUB)
    header(ws, 4, ["Linha", "Unidade"] + [f"Ano {a}" for a in ANOS])
    cell(ws, "A5", "Ano", F_FXB)
    cell(ws, "A6", "Ano calendário", F_TXT)
    for i, col in enumerate(YCOLS):
        cell(ws, f"{col}5", ANOS[i], F_FXB, FMT_ANO, align=Alignment(horizontal="center"))
        cell(ws, f"{col}6", f"={pr('ano_base', 'C')}+{col}$5", F_LK, FMT_ANO,
             align=Alignment(horizontal="center"))

    for b, (sc, nome) in enumerate(CEN):
        r0 = 8 + b * BLOCK
        build_block(ws, r0, sc, nome)

    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 16
    for col in YCOLS:
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "C7"
    return ws


def build_block(ws, r0, sc, nome):
    p = lambda k: pr(k, sc)
    rows = {}
    order = [
        ("titulo", None), ("flag", "Fluxo ativo (1 até o ano de saída)"),
        ("op", "Ano de operação"), ("idx", "Índice de reajuste acumulado"),
        ("occ", "Ocupação da ABL"), ("fcar", "Fator de carência"),
        ("h_rec", "RECEITAS"),
        ("pot", "Receita potencial de aluguel (100% ocupado)"),
        ("vac", "(-) Perda por vacância"), ("car", "(-) Carência"),
        ("outras", "(+) Outras receitas"), ("rbe", "(=) Receita bruta efetiva"),
        ("inad", "(-) Inadimplência"), ("re", "(=) Receita efetiva"),
        ("h_desp", "DESPESAS DO PROPRIETÁRIO"),
        ("adm", "(-) Taxa de administração"), ("nrep", "(-) Custos não repassáveis"),
        ("vago", "(-) Condomínio + IPTU de lojas vagas"), ("res", "(-) Reserva técnica"),
        ("noi", "(=) NOI (resultado operacional líquido)"), ("mnoi", "Margem NOI / receita efetiva"),
        ("trib", "(-) Tributos sobre a receita"),
        ("h_capex", "INVESTIMENTO (CAPEX)"),
        ("terr", "Terreno + ITBI/escritura"), ("proj", "Projetos e aprovações"),
        ("taxas", "Taxas, licenças e outorga"), ("obra", "Obra civil"),
        ("infra", "Infraestrutura externa"), ("ger", "Gerenciamento de obra"),
        ("cont", "Contingência"), ("ti", "Contribuição a lojistas (TI)"),
        ("com", "Comissão de comercialização"), ("mkt", "Marketing de lançamento"),
        ("capex", "(=) CAPEX total"),
        ("h_saida", "SAÍDA"),
        ("venda", "Valor de venda (NOI do ano seguinte / cap rate)"),
        ("cvenda", "(-) Custos de venda"), ("ir", "(-) IR/CSLL sobre ganho de capital"),
        ("fcl", "(=) FLUXO DE CAIXA LIVRE DO PROJETO"),
        ("acum", "Fluxo acumulado"), ("acum_op", "Fluxo acumulado sem a venda"),
        ("blank", None),
        ("h_ind", "INDICADORES DO CENÁRIO"),
        ("i_custo", "Custo total do projeto"),
        ("i_noi", "NOI estabilizado (Ano 4, a preços do Ano 0)"),
        ("i_yoc", "Yield on Cost (NOI estabilizado / custo total)"),
        ("i_valor", "Valor do ativo estabilizado (NOI / cap rate, preços Ano 0)"),
        ("i_criacao", "Criação de valor (valor - custo)"),
        ("i_margem", "Margem de desenvolvimento (valor / custo - 1)"),
        ("i_payback", "Payback simples (custo / NOI estabilizado)"),
        ("i_ano_pb", "Ano de recuperação do capital (sem venda)"),
        ("i_tir", "TIR do projeto (nominal, após tributos)"),
        ("i_vpl", "VPL à TMA"),
        ("i_moic", "Múltiplo sobre o capital investido (MOIC)"),
        ("i_expo", "Exposição máxima de caixa"),
        ("i_alug_min", "Aluguel médio mínimo p/ atingir YoC alvo (preços Ano 0)"),
        ("i_obra_max", "Custo de obra máximo p/ atingir YoC alvo"),
        ("i_k", "Aux.: sensibilidade do custo total a R$ 1/m² de obra"),
        ("i_b", "Aux.: custo anual de vagas estabilizado (preços Ano 0)"),
        ("i_a", "Aux.: NOI por R$ 1/m² de aluguel efetivo (preços Ano 0)"),
        ("i_cc", "Aux.: comissão por R$ 1/m² de aluguel do Mix"),
    ]
    for i, (k, _) in enumerate(order):
        rows[k] = r0 + i
        F[(sc, k)] = r0 + i
    R = rows

    # título do bloco
    for c in range(1, COL0 + len(ANOS)):
        ws.cell(row=r0, column=c).fill = FILL_HDR
    cell(ws, f"A{r0}", f"CENÁRIO {nome.upper()}", F_HDR, fill=FILL_HDR)

    units = {
        "flag": "0/1", "op": "ano", "idx": "índice", "occ": "% ABL", "fcar": "fator",
        "mnoi": "%",
    }
    fmts = {"flag": FMT_ANO, "op": FMT_ANO, "idx": FMT_IDX, "occ": FMT_PCT, "fcar": FMT_DEC,
            "mnoi": FMT_PCT}
    bold = {"rbe", "re", "noi", "capex", "fcl"}

    for k, label in order[1:]:
        r = R[k]
        if label is None:
            continue
        if k.startswith("h_"):
            section(ws, r, label, COL0 + len(ANOS) - 1)
            continue
        cell(ws, f"A{r}", label, F_FXB if k in bold else F_TXT)
        if k.startswith("i_"):
            continue
        cell(ws, f"B{r}", units.get(k, "R$"), F_NOTE)

    abl, alug = M["abl"], M["aluguel"]
    ac, ext = M["ac"], M["ext"]
    for i, L in enumerate(YCOLS):
        y = f"{L}$5"
        nxt = YCOLS[i + 1] if i + 1 < len(YCOLS) else None
        prv = YCOLS[i - 1] if i > 0 else None
        f = {}
        f["flag"] = f"=IF({y}<={p('ano_saida')},1,0)"
        f["op"] = f"=MAX(0,{y}-1)"
        f["idx"] = f"=(1+{p('reajuste')})^{y}"
        op = f"{L}{R['op']}"
        occ = f"{L}{R['occ']}"
        idx = f"{L}{R['idx']}"
        f["occ"] = (f"=IF({op}=0,0,IF({op}=1,{p('occ1')},IF({op}=2,{p('occ2')},"
                    f"1-{p('vacancia')})))")
        f["fcar"] = f"=IF({op}=1,MAX(0,12-{p('carencia')})/12,1)"
        pot = f"{L}{R['pot']}"
        f["pot"] = f"=IF({op}>0,{abl}*{alug}*{p('fator_alug')}*12*{idx},0)"
        f["vac"] = f"=-{pot}*(1-{occ})"
        f["car"] = f"=-{pot}*{occ}*(1-{L}{R['fcar']})"
        f["outras"] = f"=({pot}+{L}{R['vac']}+{L}{R['car']})*{p('outras')}"
        f["rbe"] = f"=SUM({L}{R['pot']}:{L}{R['outras']})"
        f["inad"] = f"=-{L}{R['rbe']}*{p('inad')}"
        f["re"] = f"={L}{R['rbe']}+{L}{R['inad']}"
        re_ = f"{L}{R['re']}"
        f["adm"] = f"=-{re_}*{p('adm')}"
        f["nrep"] = f"=-{re_}*{p('nrep')}"
        f["vago"] = f"=-IF({op}>0,{abl}*(1-{occ})*{p('custo_vago')}*12*{idx},0)"
        f["res"] = f"=-{re_}*{p('reserva')}"
        f["noi"] = f"={re_}+SUM({L}{R['adm']}:{L}{R['res']})"
        f["mnoi"] = f"=IFERROR({L}{R['noi']}/{re_},0)"
        f["trib"] = f"=-{re_}*{p('tributos')}*{L}{R['flag']}"
        f["terr"] = f"=IF({y}=0,-{p('terreno_rs')}*(1+{p('itbi')}),0)"
        f["proj"] = f"=IF({y}=0,-{p('projetos')}*{ac}*{p('obra_m2')},0)"
        f["taxas"] = f"=IF({y}=0,-{p('taxas')},0)"
        f["obra"] = (f"=IF({y}=0,-{ac}*{p('obra_m2')}*{p('p_obra_a0')},"
                     f"IF({y}=1,-{ac}*{p('obra_m2')}*(1-{p('p_obra_a0')})*(1+{p('incc')}),0))")
        f["infra"] = f"=IF({y}=1,-{ext}*{p('infra_m2')}*(1+{p('incc')}),0)"
        f["ger"] = f"={L}{R['obra']}*{p('gerenc')}"
        f["cont"] = f"=({L}{R['obra']}+{L}{R['infra']})*{p('conting')}"
        f["ti"] = f"=IF({y}=1,-{p('ti_m2')}*{abl},0)"
        f["com"] = (f"=IF({y}=1,-{p('comissao')}*{abl}*{alug}*{p('fator_alug')}"
                    f"*(1-{p('vacancia')})*(1+{p('reajuste')}),0)")
        f["mkt"] = f"=IF({y}=1,-{p('marketing')},0)"
        f["capex"] = f"=SUM({L}{R['terr']}:{L}{R['mkt']})"
        noi_fwd = (f"{nxt}{R['noi']}" if nxt else
                   f"{L}{R['noi']}*(1+{p('reajuste')})")
        venda = f"{L}{R['venda']}"
        f["venda"] = f"=IF({y}={p('ano_saida')},{noi_fwd}/{p('cap_saida')},0)"
        f["cvenda"] = f"=-{venda}*{p('custo_venda')}"
        capex_all = f"${FIRST}${R['capex']}:${LAST}${R['capex']}"
        f["ir"] = (f"=IF({venda}>0,-MAX(0,{venda}+{L}{R['cvenda']}+SUM({capex_all}))"
                   f"*{p('ir_ganho')},0)")
        f["fcl"] = (f"=({L}{R['noi']}+{L}{R['trib']}+{L}{R['capex']}+{venda}"
                    f"+{L}{R['cvenda']}+{L}{R['ir']})*{L}{R['flag']}")
        fcl = f"{L}{R['fcl']}"
        f["acum"] = f"={prv}{R['acum']}+{fcl}" if prv else f"={fcl}"
        sale = f"({venda}+{L}{R['cvenda']}+{L}{R['ir']})"
        f["acum_op"] = (f"={prv}{R['acum_op']}+{fcl}-{sale}" if prv else f"={fcl}-{sale}")

        for k, formula in f.items():
            fmt = fmts.get(k, FMT_RS)
            font = F_FXB if k in bold else F_FX
            c = cell(ws, f"{L}{R[k]}", formula, font, fmt)
            if k in bold:
                c.fill = FILL_TOT
        for k in bold:
            ws[f"A{R[k]}"].fill = FILL_TOT
            ws[f"B{R[k]}"].fill = FILL_TOT

    # ---- indicadores (coluna C)
    rngf = f"{FIRST}{R['fcl']}:{LAST}{R['fcl']}"
    g = YCOLS[4]  # Ano 4 = 1º ano estabilizado
    custo = f"$C${R['i_custo']}"
    noi_e = f"$C${R['i_noi']}"
    alvo = f"({p('cap_saida')}+{p('spread')})"
    k_ = f"$C${R['i_k']}"
    a_ = f"$C${R['i_a']}"
    b_ = f"$C${R['i_b']}"
    cc_ = f"$C${R['i_cc']}"
    ind = {
        "i_custo": (f"=-SUM({FIRST}{R['capex']}:{LAST}{R['capex']})", FMT_RS),
        "i_noi": (f"=IFERROR({g}{R['noi']}/{g}{R['idx']},0)", FMT_RS),
        "i_yoc": (f"=IFERROR({noi_e}/{custo},0)", FMT_PCT2),
        "i_valor": (f"=IFERROR({noi_e}/{p('cap_saida')},0)", FMT_RS),
        "i_criacao": (f"=$C${R['i_valor']}-{custo}", FMT_RS),
        "i_margem": (f"=IFERROR($C${R['i_valor']}/{custo}-1,0)", FMT_PCT),
        "i_payback": (f"=IFERROR({custo}/{noi_e},0)", FMT_ANOS),
        "i_ano_pb": (f'=IF(COUNTIF({FIRST}{R["acum_op"]}:{YCOLS[-2]}{R["acum_op"]},"<0")'
                     f'>{p("ano_saida")},"Não recupera",'
                     f'COUNTIF({FIRST}{R["acum_op"]}:{YCOLS[-2]}{R["acum_op"]},"<0"))',
                     '"Ano "0'),
        "i_tir": (f'=IFERROR(IRR({rngf}),"n/a")', FMT_PCT),
        "i_vpl": (f"={FIRST}{R['fcl']}+NPV({p('tma')},{YCOLS[1]}{R['fcl']}:{LAST}{R['fcl']})",
                  FMT_RS),
        "i_moic": (f'=IFERROR(SUMIF({rngf},">0")/-SUMIF({rngf},"<0"),0)', FMT_X),
        "i_expo": (f"=-MIN({FIRST}{R['acum']}:{LAST}{R['acum']})", FMT_RS),
        "i_alug_min": (f"=IFERROR(({alvo}*({custo}-{cc_}*{alug})+{b_})"
                       f"/({a_}*{p('fator_alug')}-{alvo}*{cc_}),0)", FMT_RS2),
        "i_obra_max": (f"=IFERROR(({noi_e}/{alvo}-({custo}-{k_}*{p('obra_m2')}))/{k_},0)",
                       FMT_RS),
        "i_k": (f"={ac}*(({p('p_obra_a0')}+(1-{p('p_obra_a0')})*(1+{p('incc')}))"
                f"*(1+{p('gerenc')}+{p('conting')})+{p('projetos')})", FMT_RS),
        "i_b": (f"={abl}*{p('vacancia')}*{p('custo_vago')}*12", FMT_RS),
        "i_cc": (f"={p('comissao')}*{abl}*{p('fator_alug')}*(1-{p('vacancia')})"
                 f"*(1+{p('reajuste')})", FMT_RS),
        "i_a": (f"={abl}*12*(1-{p('vacancia')})*(1+{p('outras')})*(1-{p('inad')})"
                f"*(1-{p('adm')}-{p('nrep')}-{p('reserva')})", FMT_RS),
    }
    notes = {
        "i_alug_min": "R$/m²/mês sobre o aluguel do Mix (antes do fator do cenário). "
                      "YoC alvo = cap rate de saída + spread mínimo.",
        "i_obra_max": "R$/m² de área construída. YoC alvo = cap rate de saída + spread.",
        "i_ano_pb": "Primeiro ano com fluxo acumulado (sem venda) >= 0.",
    }
    for k, (formula, fmt) in ind.items():
        c = cell(ws, f"C{R[k]}", formula, F_FXB if k in ("i_tir", "i_yoc", "i_vpl") else F_FX,
                 fmt)
        if k in notes:
            cell(ws, f"D{R[k]}", notes[k], F_NOTE)
        if k in ("i_k", "i_b", "i_a", "i_cc"):
            ws[f"A{R[k]}"].font = F_NOTE


# ================================================================= RESUMO
def build_resumo(wb):
    ws = wb.create_sheet("Resumo", 1)
    cell(ws, "A1", "RESUMO EXECUTIVO - VIABILIDADE DE STRIP MALL", F_TITULO)
    cell(ws, "A2", f"=Premissas!C{P['nome']}&\"  |  \"&Premissas!C{P['local']}"
         f"&\"  |  Ano-base \"&Premissas!C{P['ano_base']}", F_LKB)
    cell(ws, "A3", "Todos os valores são calculados a partir das abas Premissas, Mix e Fluxo. "
         "Não edite esta aba.", F_SUB)
    header(ws, 5, ["Indicador", "Pessimista", "Realista", "Otimista", "Leitura"])

    def link(key, sc):
        return f"=Fluxo!$C${F[(sc, key)]}"

    lines = [
        ("SEC", "PRODUTO"),
        ("ABL total (m²)", lambda sc: f"={M['abl']}", FMT_NUM, ""),
        ("Área construída (m²)", lambda sc: f"={M['ac']}", FMT_NUM, ""),
        ("Nº de lojas", lambda sc: f"={M['lojas']}", FMT_NUM, ""),
        ("Aluguel médio efetivo (R$/m²/mês, preços Ano 0)",
         lambda sc: f"={M['aluguel']}*{pr('fator_alug', sc)}", FMT_RS2,
         "Aluguel do Mix x fator do cenário"),
        ("Vacância estabilizada", lambda sc: f"={pr('vacancia', sc)}", FMT_PCT, ""),
        ("SEC", "INVESTIMENTO"),
        ("Custo total do projeto (R$)", lambda sc: link("i_custo", sc), FMT_RS,
         "Terreno + obra + soft costs + contingência"),
        ("Custo total por m² de ABL (R$/m²)",
         lambda sc: f"=IFERROR(Fluxo!$C${F[(sc, 'i_custo')]}/{M['abl']},0)", FMT_RS, ""),
        ("Exposição máxima de caixa (R$)", lambda sc: link("i_expo", sc), FMT_RS,
         "Capital próprio necessário sem financiamento"),
        ("SEC", "OPERAÇÃO ESTABILIZADA (Ano 4, preços do Ano 0)"),
        ("NOI anual estabilizado (R$)", lambda sc: link("i_noi", sc), FMT_RS, ""),
        ("NOI mensal estabilizado (R$)",
         lambda sc: f"=Fluxo!$C${F[(sc, 'i_noi')]}/12", FMT_RS, ""),
        ("Margem NOI / receita efetiva",
         lambda sc: f"=Fluxo!${YCOLS[4]}${F[(sc, 'mnoi')]}", FMT_PCT, ""),
        ("SEC", "RETORNO"),
        ("Yield on Cost", lambda sc: link("i_yoc", sc), FMT_PCT2,
         "Deve superar o cap rate de saída + spread"),
        ("Cap rate de saída", lambda sc: f"={pr('cap_saida', sc)}", FMT_PCT2, "Hipótese"),
        ("Spread de desenvolvimento (YoC - cap rate)",
         lambda sc: f"=Fluxo!$C${F[(sc, 'i_yoc')]}-{pr('cap_saida', sc)}", FMT_PCT2,
         "Mínimo exigido: ver Premissas"),
        ("Valor do ativo estabilizado (R$)", lambda sc: link("i_valor", sc), FMT_RS, ""),
        ("Criação de valor (R$)", lambda sc: link("i_criacao", sc), FMT_RS, ""),
        ("Margem de desenvolvimento", lambda sc: link("i_margem", sc), FMT_PCT, ""),
        ("Payback simples", lambda sc: link("i_payback", sc), FMT_ANOS, "Custo / NOI"),
        ("Ano de recuperação do capital (sem venda)", lambda sc: link("i_ano_pb", sc),
         '"Ano "0', ""),
        ("TIR do projeto (nominal, após tributos)", lambda sc: link("i_tir", sc), FMT_PCT,
         "Comparar com a TMA"),
        ("TMA", lambda sc: f"={pr('tma', sc)}", FMT_PCT, ""),
        ("VPL à TMA (R$)", lambda sc: link("i_vpl", sc), FMT_RS, "> 0 cria valor"),
        ("MOIC", lambda sc: link("i_moic", sc), FMT_X, ""),
        ("SEC", "LIMITES DE VIABILIDADE (para YoC = cap rate + spread)"),
        ("Aluguel médio mínimo do Mix (R$/m²/mês)", lambda sc: link("i_alug_min", sc),
         FMT_RS2, "Comparar com o aluguel do Mix"),
        ("Aluguel médio atual do Mix (R$/m²/mês)", lambda sc: f"={M['aluguel']}", FMT_RS2,
         ""),
        ("Custo de obra máximo (R$/m² AC)", lambda sc: link("i_obra_max", sc), FMT_RS,
         "Comparar com o orçamento"),
        ("Custo de obra adotado (R$/m² AC)", lambda sc: f"={pr('obra_m2', sc)}", FMT_RS, ""),
    ]
    r = 6
    R = {}
    for item in lines:
        if item[0] == "SEC":
            section(ws, r, item[1], 5)
            r += 1
            continue
        lab, fn, fmt, note = item
        R[lab] = r
        cell(ws, f"A{r}", lab, F_TXT)
        for sc, _ in CEN:
            f = fn(sc)
            cell(ws, f"{sc}{r}", f, F_LK, fmt, align=Alignment(horizontal="right"))
        cell(ws, f"E{r}", note, F_NOTE)
        r += 1

    # veredito
    section(ws, r, "VEREDITO", 5)
    r += 1
    rv = r
    cell(ws, f"A{r}", "Parecer por cenário", F_FXB)
    tir_r, yoc_r = R["TIR do projeto (nominal, após tributos)"], R["Yield on Cost"]
    for sc, _ in CEN:
        tir = f"IF(ISNUMBER({sc}{tir_r}),{sc}{tir_r},-1)"
        alvo = f"({pr('cap_saida', sc)}+{pr('spread', sc)})"
        f = (f'=IF(AND({tir}>={pr("tma", sc)},{sc}{yoc_r}>={alvo}),"VIÁVEL",'
             f'IF(OR({tir}>={pr("tma", sc)},{sc}{yoc_r}>={alvo}),"ATENÇÃO","INVIÁVEL"))')
        cell(ws, f"{sc}{r}", f, F_FXB, align=Alignment(horizontal="center"))
    cell(ws, f"E{r}", "VIÁVEL = TIR >= TMA e YoC >= cap + spread | ATENÇÃO = só um critério "
         "| INVIÁVEL = nenhum", F_NOTE)
    rng = f"B{rv}:D{rv}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"VIÁVEL"'],
                                  fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"ATENÇÃO"'],
                                  fill=PatternFill("solid", fgColor="FFEB9C")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"INVIÁVEL"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    r += 1
    cell(ws, f"A{r}", "Checagens de produto com ALERTA (aba Mix)", F_TXT)
    first, last = M["checks_first"], M["checks_last"]
    cell(ws, f"C{r}", f'=COUNTIF(Mix!C{first}:C{last},"ALERTA")', F_LKB, FMT_ANO,
         align=Alignment(horizontal="right"))
    cell(ws, f"E{r}", "Qualquer ALERTA exige revisão do produto antes de seguir", F_NOTE)

    # gráfico de FCL
    r += 2
    cell(ws, f"A{r}", "Fluxo de caixa livre por cenário (R$) - dados da aba Fluxo", F_SEC)
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.title = "Fluxo de caixa livre do projeto por ano"
    ch.y_axis.title = "R$"
    ch.y_axis.numFmt = "#,##0"
    ch.height, ch.width = 8, 22
    fl = wb["Fluxo"]
    for sc, nome in CEN:
        row = F[(sc, "fcl")]
        data = Reference(fl, min_col=COL0, max_col=COL0 + 11, min_row=row, max_row=row)
        ch.add_data(data, from_rows=True, titles_from_data=False)
    for s, (_, nome) in zip(ch.series, CEN):
        s.tx = SeriesLabel(v=nome)
    cats = Reference(fl, min_col=COL0, max_col=COL0 + 11, min_row=4, max_row=4)
    ch.set_categories(cats)
    ws.add_chart(ch, f"A{r + 1}")

    ws.column_dimensions["A"].width = 52
    for col in "BCD":
        ws.column_dimensions[col].width = 18
    ws.column_dimensions["E"].width = 70
    ws.freeze_panes = "B6"
    return ws


# ========================================================== SENSIBILIDADE
def build_sens(wb):
    ws = wb.create_sheet("Sensibilidade")
    sc = "C"
    cell(ws, "A1", "SENSIBILIDADE - CENÁRIO REALISTA (operação estabilizada)", F_TITULO)
    cell(ws, "A2", "Variações em azul são editáveis. Cálculo fechado sobre o NOI "
         "estabilizado (preços do Ano 0); não altera o fluxo de caixa.", F_SUB)
    noi = f"Fluxo!$C${F[(sc, 'i_noi')]}"
    custo = f"Fluxo!$C${F[(sc, 'i_custo')]}"
    k = f"Fluxo!$C${F[(sc, 'i_k')]}"
    b = f"Fluxo!$C${F[(sc, 'i_b')]}"
    cc = f"Fluxo!$C${F[(sc, 'i_cc')]}*{M['aluguel']}"
    obra = pr("obra_m2", sc)
    cap = pr("cap_saida", sc)
    var = [-0.20, -0.10, 0.0, 0.10, 0.20]
    capv = [-0.01, -0.005, 0.0, 0.005, 0.01]

    # Tabela 1: YoC
    r = 4
    section(ws, r, "1. YIELD ON COST: variação do aluguel (linhas) x variação do custo de obra "
            "(colunas)", 7)
    cell(ws, f"A{r + 1}", "Aluguel \\ Custo de obra", F_FXB)
    cols = ["C", "D", "E", "F", "G"]
    for c, v in zip(cols, var):
        cell(ws, f"{c}{r + 1}", v, F_IN, '+0%;-0%;0%', align=Alignment(horizontal="center"))
    for i, v in enumerate(var):
        rr = r + 2 + i
        cell(ws, f"B{rr}", v, F_IN, '+0%;-0%;0%', align=Alignment(horizontal="center"))
        for c in cols:
            f = (f"=IFERROR(({noi}*(1+$B{rr})+{b}*$B{rr})/({custo}+{k}*{obra}*{c}${r + 1}+{cc}*$B{rr}),0)")
            cell(ws, f"{c}{rr}", f, F_FX, FMT_PCT2)
    t1 = f"C{r + 2}:G{r + 6}"
    ws.conditional_formatting.add(t1, ColorScaleRule(start_type="min", start_color="F8696B",
                                                     mid_type="percentile", mid_value=50,
                                                     mid_color="FFEB84", end_type="max",
                                                     end_color="63BE7B"))
    cell(ws, f"A{r + 8}", "Referência: YoC alvo (cap rate de saída + spread) =", F_TXT)
    cell(ws, f"D{r + 8}", f"={cap}+{pr('spread', sc)}", F_LKB, FMT_PCT2)

    # Tabela 2: margem de desenvolvimento
    r = 15
    section(ws, r, "2. MARGEM DE DESENVOLVIMENTO: variação do aluguel (linhas) x cap rate de "
            "saída (colunas)", 7)
    cell(ws, f"A{r + 1}", "Aluguel \\ Cap rate", F_FXB)
    for c, v in zip(cols, capv):
        cell(ws, f"{c}{r + 1}", f"={cap}+({v})", F_FX, FMT_PCT2,
             align=Alignment(horizontal="center"))
    for i, v in enumerate(var):
        rr = r + 2 + i
        cell(ws, f"B{rr}", v, F_IN, '+0%;-0%;0%', align=Alignment(horizontal="center"))
        for c in cols:
            f = f"=IFERROR(({noi}*(1+$B{rr})+{b}*$B{rr})/{c}${r + 1}/({custo}+{cc}*$B{rr})-1,0)"
            cell(ws, f"{c}{rr}", f, F_FX, FMT_PCT)
    t2 = f"C{r + 2}:G{r + 6}"
    ws.conditional_formatting.add(t2, ColorScaleRule(start_type="min", start_color="F8696B",
                                                     mid_type="num", mid_value=0,
                                                     mid_color="FFEB84", end_type="max",
                                                     end_color="63BE7B"))
    cell(ws, f"A{r + 8}", "Custo total ajustado pela variação da comissão de comercialização "
         "(função do aluguel).", F_NOTE)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 10
    for c in cols:
        ws.column_dimensions[c].width = 13
    return ws


# ============================================================= INSTRUÇÕES
def build_instrucoes(wb):
    ws = wb.active
    ws.title = "Instruções"
    cell(ws, "A1", "MODELO PADRONIZADO DE VIABILIDADE - STRIP MALL", F_TITULO)
    cell(ws, "A2", "Nexa Malls | Template reutilizável | Valores de exemplo são ilustrativos",
         F_SUB)
    txt = [
        ("COMO USAR", None),
        ("1.", "Premissas: preencha as células em AZUL. Amarelo = premissa-chave que varia por "
               "cenário (Pessimista / Realista / Otimista)."),
        ("2.", "Premissas não amarelas: o valor é digitado em Realista e replicado por fórmula "
               "nos demais cenários. Sobrescreva em Pessimista/Otimista se quiser diferenciar."),
        ("3.", "Mix: defina segmentos, nº de lojas, área média e aluguel/m². A ABL e o aluguel "
               "médio ponderado alimentam o modelo. Linhas 12-16 estão livres."),
        ("4.", "Mix > Checagens: verifique se o produto cabe no terreno (CA, TO, vagas) e se o "
               "mix tem âncora e baixa concentração."),
        ("5.", "Resumo: indicadores e parecer por cenário. Fluxo: DRE/fluxo anual completo. "
               "Sensibilidade: YoC e margem sob variações de aluguel, obra e cap rate."),
        ("LEGENDA DE CORES", None),
        ("Azul", "Input editável (valor digitado)"),
        ("Preto", "Fórmula - não editar"),
        ("Verde", "Link para outra aba"),
        ("Amarelo", "Premissa-chave por cenário"),
        ("LÓGICA DO MODELO", None),
        ("Fórmulas", "Receita potencial = ABL x aluguel/m² x fator x 12 x índice de reajuste"),
        ("", "Receita efetiva = potencial - vacância - carência + outras - inadimplência"),
        ("", "NOI = receita efetiva - administração - não repassáveis - encargos de vagas - "
             "reserva técnica (NOI antes de tributos, padrão de mercado para cap rate)"),
        ("", "Yield on Cost = NOI estabilizado / custo total | Cap rate = NOI / valor | "
             "Payback simples = custo / NOI"),
        ("", "Fluxo livre = NOI - tributos - CAPEX + venda líquida (desalavancado, nominal)"),
        ("Cronograma", "Ano 0: terreno, projetos, taxas e parte da obra | Ano 1: obra | Ano 2: "
                       "1º ano de operação | Ano 4 em diante: operação estabilizada"),
        ("Saída", "Valor de venda = NOI do ano seguinte / cap rate de saída, menos custos de "
                  "venda e IR/CSLL sobre o ganho"),
        ("Critério", "VIÁVEL se TIR >= TMA e YoC >= cap rate de saída + spread mínimo"),
        ("LIMITAÇÕES (versão 1.0)", None),
        ("-", "Sem financiamento (análise desalavancada). Alavancagem é evolução prevista."),
        ("-", "Sem aluguel percentual, CDU/luvas e reajustes escalonados por contrato."),
        ("-", "Tributação simplificada (lucro presumido). Validar estrutura com contador."),
        ("-", "Fluxo anual: não captura o descasamento mensal de obra e entrada de receita."),
        ("-", "Premissas de exemplo NÃO são dados de mercado verificados: substitua por "
              "pesquisa local (aluguel, cap rate, custo de obra) antes de decidir."),
    ]
    r = 4
    for a, b in txt:
        if b is None:
            section(ws, r, a, 2)
        else:
            cell(ws, f"A{r}", a, F_FXB)
            cell(ws, f"B{r}", b, F_TXT, align=Alignment(wrap_text=True, vertical="top"))
            if a == "Azul":
                ws[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=AZUL)
            elif a == "Verde":
                ws[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=VERDE)
            elif a == "Amarelo":
                ws[f"A{r}"].fill = FILL_KEY
        r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 120
    return ws


def main(out):
    wb = Workbook()
    build_instrucoes(wb)
    build_premissas(wb)
    build_mix(wb)
    build_fluxo(wb)
    build_resumo(wb)
    build_sens(wb)
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Modelo gerado: {out}")


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    dest = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "modelo" / "Viabilidade_Strip_Mall.xlsx"
    main(dest)
