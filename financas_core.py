# -*- coding: utf-8 -*-
"""
Núcleo do Controle Financeiro (sem interface gráfica).

Aqui ficam: leitura/gravação da planilha Excel, cálculos, conquistas e
ranking de investimentos. O app (app_financas.py) usa este módulo, e a
planilha controle_financeiro.xlsx é a "fonte da verdade": o que mudar nela
aparece no app e o que mudar no app é gravado nela.
"""
import copy
import datetime as dt
import os
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ARQUIVO_PADRAO = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "controle_financeiro.xlsx"
)

CATEGORIAS = [
    "Moradia", "Alimentação", "Transporte", "Saúde", "Educação",
    "Lazer", "Assinaturas", "Cartão/Dívidas", "Compras", "Outros",
]

INDEXADORES = ["%CDI", "%Selic", "IPCA+", "Fixo"]

# Valores de referência pesquisados em 25/10/2026 (Selic 14,00% a.a. desde
# agosto/2026, CDI ~13,90% a.a., IPCA 12m ~4,44%, poupança ~6,17% a.a.).
# As taxas reais variam por banco/corretora: edite na aba "Investimentos".
INVESTIMENTOS_PADRAO = [
    dict(nome="Tesouro Selic", indexador="%Selic", valor=100, ir=0.175,
         liquidez="Diária", risco="Baixo",
         obs="Bom para reserva de emergência. Garantido pelo governo. IR regressivo de 22,5% a 15%."),
    dict(nome="CDB liquidez diária (100% CDI)", indexador="%CDI", valor=100, ir=0.175,
         liquidez="Diária", risco="Baixo",
         obs="Procure bancos que paguem 100% do CDI ou mais. FGC cobre até R$ 250 mil por CPF por instituição."),
    dict(nome="CDB prazo fixo (110% CDI)", indexador="%CDI", valor=110, ir=0.175,
         liquidez="No vencimento", risco="Baixo/médio",
         obs="Rende mais, mas o dinheiro fica preso até o vencimento. Confira o % realmente oferecido."),
    dict(nome="LCI/LCA (90% CDI)", indexador="%CDI", valor=90, ir=0.0,
         liquidez="Carência (90 dias a 2 anos)", risco="Baixo/médio",
         obs="Isenta de IR para pessoa física. FGC cobre até R$ 250 mil."),
    dict(nome="Tesouro IPCA+ (longo prazo)", indexador="IPCA+", valor=8.12, ir=0.15,
         liquidez="Diária (com oscilação)", risco="Médio",
         obs="Protege da inflação. O preço oscila antes do vencimento: só para metas de 3+ anos. Taxa de referência, confira no Tesouro Direto."),
    dict(nome="Poupança", indexador="Fixo", valor=6.17, ir=0.0,
         liquidez="Diária", risco="Baixo",
         obs="Rende bem menos que a renda fixa com a Selic alta."),
]

# (chave, rótulo na planilha, tipo, observação)
CFG = [
    ("salario", "Salário mensal líquido (R$)", "moeda", "Quanto entra por mês, já com descontos."),
    ("pct_guardar", "% do salário para guardar por mês", "pct", "Ex.: 20% = guardar 1/5 do salário todo mês."),
    ("meta_nome", "Nome da meta financeira", "txt", "Ex.: Reserva de emergência, Viagem, Carro..."),
    ("meta_valor", "Valor da meta (R$)", "moeda", "Quanto você quer juntar no total."),
    ("prazo_meta_meses", "Prazo para a meta (meses)", "int", "Em quantos meses quer bater a meta."),
    ("saldo_inicial", "Já guardado antes de usar o app (R$)", "moeda", "Saldo que você já tinha guardado."),
    ("data_saldo_inicial", "Data desse saldo inicial", "data", "Usada para calcular o rendimento dele."),
    ("rendimento_anual", "Rendimento anual líquido do que você guarda (%)", "pct", "Estimativa depois do IR. Veja a aba Investimentos."),
    ("selic", "Selic atual (% a.a.)", "pct", "Referência para o ranking de investimentos."),
    ("cdi", "CDI atual (% a.a.)", "pct", "Referência para o ranking de investimentos."),
    ("ipca", "IPCA acumulado 12 meses (%)", "pct", "Referência para o Tesouro IPCA+."),
]
CFG_LINHA = {c[0]: i + 2 for i, c in enumerate(CFG)}

NOMES_ABAS = ["Resumo", "Config", "Gastos", "Guardado", "Investimentos", "Conquistas"]
VERDE = "1F6F4A"


# ----------------------------------------------------------------- estado
@dataclass
class Estado:
    salario: float = 0.0
    pct_guardar: float = 0.20
    meta_nome: str = "Minha meta"
    meta_valor: float = 10000.0
    prazo_meta_meses: int = 12
    saldo_inicial: float = 0.0
    data_saldo_inicial: dt.date = field(default_factory=dt.date.today)
    rendimento_anual: float = 0.115
    selic: float = 0.14
    cdi: float = 0.139
    ipca: float = 0.0444
    gastos: list = field(default_factory=list)    # {data, descricao, categoria, valor}
    aportes: list = field(default_factory=list)   # {data, valor, obs}
    investimentos: list = field(
        default_factory=lambda: copy.deepcopy(INVESTIMENTOS_PADRAO)
    )


# ------------------------------------------------------------ conversões
def brl(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-R$ " if v < 0 else "R$ ") + s


def num(v, default=0.0):
    """Converte número/texto ('1.234,56', 'R$ 50') em float."""
    if v is None or v == "" or isinstance(v, bool):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").replace("%", "").strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def pct(v, default=0.0):
    """Lê porcentagem vinda do Excel: 0,2 ou 20% ou 20 -> 0.20."""
    if v is None or v == "":
        return default
    if isinstance(v, str):
        x = num(v, None)
        if x is None:
            return default
        return x / 100 if ("%" in v or x > 1) else x
    x = float(v)
    return x / 100 if x > 1 else x


def data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


def _indexador(v):
    t = str(v or "").strip().lower().replace(" ", "")
    for ix in INDEXADORES:
        if t == ix.lower():
            return ix
    return "Fixo"


# --------------------------------------------------------- ler a planilha
def _aplicar_cfg(e, chave, tipo, val):
    if val is None or val == "":
        return
    atual = getattr(e, chave)
    if tipo == "moeda":
        setattr(e, chave, num(val, atual))
    elif tipo == "pct":
        setattr(e, chave, pct(val, atual))
    elif tipo == "int":
        setattr(e, chave, int(num(val, atual)))
    elif tipo == "data":
        d = data(val)
        if d:
            setattr(e, chave, d)
    else:
        setattr(e, chave, str(val).strip())


def carregar(caminho=ARQUIVO_PADRAO):
    wb = load_workbook(caminho, data_only=True)
    e = Estado()
    hoje = dt.date.today()

    if "Config" in wb.sheetnames:
        por_rotulo = {c[1]: c for c in CFG}
        for idx, (rot, val) in enumerate(
            wb["Config"].iter_rows(min_row=2, max_col=2, values_only=True)
        ):
            cfg = por_rotulo.get(str(rot).strip()) if rot else None
            if cfg is None and idx < len(CFG):
                cfg = CFG[idx]
            if cfg is not None:
                _aplicar_cfg(e, cfg[0], cfg[2], val)

    if "Gastos" in wb.sheetnames:
        for d, desc, cat, val in wb["Gastos"].iter_rows(
            min_row=2, max_col=4, values_only=True
        ):
            v = num(val, None)
            if v is None:
                continue
            e.gastos.append(dict(
                data=data(d) or hoje,
                descricao=str(desc or "").strip(),
                categoria=str(cat or "Outros").strip() or "Outros",
                valor=v,
            ))

    if "Guardado" in wb.sheetnames:
        for d, val, obs in wb["Guardado"].iter_rows(
            min_row=2, max_col=3, values_only=True
        ):
            v = num(val, None)
            if v is None:
                continue
            e.aportes.append(dict(
                data=data(d) or hoje, valor=v, obs=str(obs or "").strip()
            ))

    if "Investimentos" in wb.sheetnames:
        invs = []
        for row in wb["Investimentos"].iter_rows(min_row=2, max_col=7, values_only=True):
            nome, ix, val, ir, liq, risco, obs = row
            if not nome or num(val, None) is None:
                continue
            invs.append(dict(
                nome=str(nome).strip(), indexador=_indexador(ix), valor=num(val),
                ir=pct(ir, 0.0), liquidez=str(liq or ""), risco=str(risco or ""),
                obs=str(obs or ""),
            ))
        if invs:
            e.investimentos = invs
    return e


# ------------------------------------------------------ gravar a planilha
def _cabecalho(ws, titulos, larguras):
    for i, t in enumerate(titulos, 1):
        c = ws.cell(1, i, t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=VERDE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = larguras[i - 1]
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32


MOEDA = '"R$" #,##0.00'


def _aba_config(wb, e, idx):
    ws = wb.create_sheet("Config", idx)
    _cabecalho(ws, ["Parâmetro", "Valor", "Observação"], [52, 22, 60])
    for i, (chave, rotulo, tipo, obs) in enumerate(CFG, 2):
        ws.cell(i, 1, rotulo)
        c = ws.cell(i, 2, getattr(e, chave))
        c.alignment = Alignment(horizontal="right")
        c.fill = PatternFill("solid", fgColor="FFF8DC")
        c.number_format = {"moeda": MOEDA, "pct": "0.00%", "data": "dd/mm/yyyy", "int": "0"}.get(tipo, "General")
        ws.cell(i, 3, obs).font = Font(italic=True, color="666666")


def _aba_gastos(wb, e, idx):
    ws = wb.create_sheet("Gastos", idx)
    _cabecalho(ws, ["Data", "Descrição", "Categoria", "Valor (R$)"], [14, 38, 20, 16])
    for i, g in enumerate(sorted(e.gastos, key=lambda x: x["data"]), 2):
        ws.cell(i, 1, g["data"]).number_format = "dd/mm/yyyy"
        ws.cell(i, 2, g["descricao"])
        ws.cell(i, 3, g["categoria"])
        ws.cell(i, 4, g["valor"]).number_format = MOEDA
    for r in range(len(e.gastos) + 2, 1002):
        ws.cell(r, 1).number_format = "dd/mm/yyyy"
        ws.cell(r, 4).number_format = MOEDA
    dv = DataValidation(type="list", formula1='"' + ",".join(CATEGORIAS) + '"',
                        allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("C2:C1001")


def _aba_guardado(wb, e, idx):
    ws = wb.create_sheet("Guardado", idx)
    _cabecalho(ws, ["Data", "Valor guardado (R$)", "Observação"], [14, 22, 40])
    for i, a in enumerate(sorted(e.aportes, key=lambda x: x["data"]), 2):
        ws.cell(i, 1, a["data"]).number_format = "dd/mm/yyyy"
        ws.cell(i, 2, a["valor"]).number_format = MOEDA
        ws.cell(i, 3, a["obs"])
    for r in range(len(e.aportes) + 2, 1002):
        ws.cell(r, 1).number_format = "dd/mm/yyyy"
        ws.cell(r, 2).number_format = MOEDA


def _aba_investimentos(wb, e, idx):
    ws = wb.create_sheet("Investimentos", idx)
    _cabecalho(
        ws,
        ["Investimento", "Indexador (%CDI, %Selic, IPCA+ ou Fixo)",
         "Valor (% do CDI/Selic, taxa real % ou taxa % a.a.)", "IR estimado",
         "Liquidez", "Risco", "Observação"],
        [34, 24, 30, 13, 26, 14, 80],
    )
    ws.row_dimensions[1].height = 48
    for i, v in enumerate(e.investimentos, 2):
        ws.cell(i, 1, v["nome"])
        ws.cell(i, 2, v["indexador"])
        ws.cell(i, 3, v["valor"])
        ws.cell(i, 4, v["ir"]).number_format = "0.0%"
        ws.cell(i, 5, v["liquidez"])
        ws.cell(i, 6, v["risco"])
        ws.cell(i, 7, v["obs"])
    dv = DataValidation(type="list", formula1='"' + ",".join(INDEXADORES) + '"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B2:B60")


def _aba_conquistas(wb, e, idx):
    ws = wb.create_sheet("Conquistas", idx)
    _cabecalho(ws, ["Grupo", "Conquista", "Como desbloquear", "Progresso", "Status"],
               [16, 34, 56, 14, 18])
    for i, c in enumerate(conquistas(e), 2):
        ws.cell(i, 1, c["grupo"])
        ws.cell(i, 2, c["nome"])
        ws.cell(i, 3, c["desc"])
        ws.cell(i, 4, c["prog"]).number_format = "0%"
        s = ws.cell(i, 5, "★ Desbloqueada" if c["ok"] else "Bloqueada")
        s.font = Font(bold=c["ok"], color="1F6F4A" if c["ok"] else "888888")


def _aba_resumo(wb, e, idx):
    ws = wb.create_sheet("Resumo", idx)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 20
    ws["A1"] = "Controle Financeiro — resumo do mês atual"
    ws["A1"].font = Font(bold=True, size=15, color=VERDE)
    ws["A2"] = "Esta aba é calculada pelo Excel. Edite as abas Config, Gastos, Guardado e Investimentos."
    ws["A2"].font = Font(italic=True, color="666666")

    r = CFG_LINHA
    mes = 'Gastos!$A:$A,">="&$B$4,Gastos!$A:$A,"<"&$B$5'
    mes_g = 'Guardado!$A:$A,">="&$B$4,Guardado!$A:$A,"<"&$B$5'
    linhas = [
        (3, "Hoje", "=TODAY()", "dd/mm/yyyy"),
        (4, "Início do mês", "=DATE(YEAR(B3),MONTH(B3),1)", "dd/mm/yyyy"),
        (5, "Início do próximo mês", "=EDATE(B4,1)", "dd/mm/yyyy"),
        (7, "Salário", f"=Config!B{r['salario']}", MOEDA),
        (8, "Guardar no mês (meta mensal)", f"=B7*Config!B{r['pct_guardar']}", MOEDA),
        (9, "Gastos do mês", f"=SUMIFS(Gastos!$D:$D,{mes})", MOEDA),
        (10, "Ainda posso gastar", "=B7-B8-B9", MOEDA),
        (11, "Já guardado no mês", f"=SUMIFS(Guardado!$B:$B,{mes_g})", MOEDA),
        (12, "Falta guardar no mês", "=MAX(0,B8-B11)", MOEDA),
        (13, "Total guardado (sem rendimento)", f"=Config!B{r['saldo_inicial']}+SUM(Guardado!$B:$B)", MOEDA),
        (14, "Meta financeira", f"=Config!B{r['meta_valor']}", MOEDA),
        (15, "Falta para a meta (sem rendimento)", "=MAX(0,B14-B13)", MOEDA),
        (16, "% da meta", "=IF(B14=0,0,MIN(1,B13/B14))", "0.0%"),
    ]
    for lin, rot, formula, fmt in linhas:
        ws.cell(lin, 1, rot).font = Font(bold=lin in (10, 12, 15))
        c = ws.cell(lin, 2, formula)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    for lin in (3, 4, 5):
        ws.cell(lin, 1).font = Font(color="888888")
        ws.cell(lin, 2).font = Font(color="888888")

    ws["A18"] = "Gastos do mês por categoria"
    ws["A18"].font = Font(bold=True, color=VERDE)
    ws["A19"], ws["B19"] = "Categoria", "Gasto"
    for c in (ws["A19"], ws["B19"]):
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=VERDE)
    n = len(CATEGORIAS)
    for i, cat in enumerate(CATEGORIAS):
        lin = 20 + i
        ws.cell(lin, 1, cat)
        c = ws.cell(lin, 2, f"=SUMIFS(Gastos!$D:$D,Gastos!$C:$C,A{lin},{mes})")
        c.number_format = MOEDA
    ws.cell(20 + n, 1, "Outras categorias")
    c = ws.cell(20 + n, 2, f"=B9-SUM(B20:B{19 + n})")
    c.number_format = MOEDA

    pie = PieChart()
    pie.title = "Para onde foi o dinheiro (mês atual)"
    pie.add_data(Reference(ws, min_col=2, min_row=20, max_row=20 + n), titles_from_data=False)
    pie.set_categories(Reference(ws, min_col=1, min_row=20, max_row=20 + n))
    pie.height, pie.width = 8.5, 13
    ws.add_chart(pie, "D3")

    bar = BarChart()
    bar.type = "bar"
    bar.title = "Salário x guardar x gastos x sobra"
    bar.add_data(Reference(ws, min_col=2, min_row=7, max_row=10), titles_from_data=False)
    bar.set_categories(Reference(ws, min_col=1, min_row=7, max_row=10))
    bar.legend = None
    bar.height, bar.width = 7.5, 13
    ws.add_chart(bar, "D21")


def salvar(e, caminho=ARQUIVO_PADRAO):
    """Grava o estado na planilha. Levanta PermissionError se o Excel a trava."""
    if os.path.exists(caminho):
        wb = load_workbook(caminho)  # preserva abas extras do usuário
    else:
        wb = Workbook()
        wb.remove(wb.active)
    for nome in NOMES_ABAS:
        if nome in wb.sheetnames:
            del wb[nome]
    _aba_resumo(wb, e, 0)
    _aba_config(wb, e, 1)
    _aba_gastos(wb, e, 2)
    _aba_guardado(wb, e, 3)
    _aba_investimentos(wb, e, 4)
    _aba_conquistas(wb, e, 5)
    wb.active = 0
    tmp = caminho + ".tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, caminho)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


# ---------------------------------------------------------------- cálculos
def _meses_entre(ini, fim):
    y, m, out = ini.year, ini.month, []
    while (y, m) <= (fim.year, fim.month):
        out.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def _fim_mes(y, m):
    prox = dt.date(y + (m == 12), m % 12 + 1, 1)
    return prox - dt.timedelta(days=1)


def rendimento_ate(e, ate):
    """Rendimento acumulado (juros compostos pela taxa anual líquida)."""
    r = e.rendimento_anual
    itens = [(e.data_saldo_inicial, e.saldo_inicial)] + [
        (a["data"], a["valor"]) for a in e.aportes
    ]
    total = 0.0
    for d, v in itens:
        dias = (ate - d).days
        if v > 0 and dias > 0:
            total += v * ((1 + r) ** (dias / 365) - 1)
    return total


def media_gastos(e, hoje=None):
    hoje = hoje or dt.date.today()
    por_mes = {}
    for g in e.gastos:
        k = (g["data"].year, g["data"].month)
        por_mes[k] = por_mes.get(k, 0) + g["valor"]
    ult = sorted(por_mes)[-3:]
    return sum(por_mes[k] for k in ult) / len(ult) if ult else 0.0


def resumo(e, ano, mes, hoje=None):
    hoje = hoje or dt.date.today()
    por_cat, gastos_mes = {}, 0.0
    for g in e.gastos:
        if (g["data"].year, g["data"].month) == (ano, mes):
            gastos_mes += g["valor"]
            por_cat[g["categoria"]] = por_cat.get(g["categoria"], 0) + g["valor"]
    guardar_mes = e.salario * e.pct_guardar
    guardado_mes = sum(
        a["valor"] for a in e.aportes if (a["data"].year, a["data"].month) == (ano, mes)
    )
    total = e.saldo_inicial + sum(a["valor"] for a in e.aportes)
    rend = rendimento_ate(e, hoje)
    patr = total + rend
    meta = e.meta_valor
    falta = max(0.0, meta - patr)
    m = (1 + e.rendimento_anual) ** (1 / 12) - 1

    # meses até a meta no ritmo atual
    meses = None
    if patr >= meta:
        meses = 0
    elif guardar_mes > 0 or m > 0:
        saldo, n = patr, 0
        while saldo < meta and n < 1200:
            saldo = saldo * (1 + m) + guardar_mes
            n += 1
        meses = n if n < 1200 else None

    # quanto guardar por mês para cumprir o prazo
    n = max(1, e.prazo_meta_meses)
    restante = meta - patr * (1 + m) ** n
    if restante <= 0:
        pmt = 0.0
    elif m > 0:
        pmt = restante * m / ((1 + m) ** n - 1)
    else:
        pmt = restante / n

    # evolução mês a mês (guardado x guardado com rendimento)
    inicio = min([e.data_saldo_inicial] + [a["data"] for a in e.aportes] + [hoje])
    evolucao = []
    for y, mm in _meses_entre(inicio, hoje):
        fim = min(_fim_mes(y, mm), hoje)
        dep = (e.saldo_inicial if e.data_saldo_inicial <= fim else 0) + sum(
            a["valor"] for a in e.aportes if a["data"] <= fim
        )
        evolucao.append((f"{mm:02d}/{y % 100:02d}", dep, dep + rendimento_ate(e, fim)))

    return dict(
        salario=e.salario, guardar_mes=guardar_mes, gastos_mes=gastos_mes,
        pode_gastar=e.salario - guardar_mes - gastos_mes,
        guardado_mes=guardado_mes, falta_guardar_mes=max(0.0, guardar_mes - guardado_mes),
        total_guardado=total, rendimento=rend, patrimonio=patr,
        meta_valor=meta, falta_meta=falta,
        pct_meta=min(1.0, patr / meta) if meta > 0 else 0.0,
        meses_para_meta=meses, guardar_para_prazo=pmt,
        por_categoria=dict(sorted(por_cat.items(), key=lambda kv: -kv[1])),
        evolucao=evolucao,
    )


# -------------------------------------------------------------- conquistas
def _sequencia_guardando(e, hoje):
    meta_mes = e.salario * e.pct_guardar
    if meta_mes <= 0 or not e.aportes:
        return 0
    por_mes = {}
    for a in e.aportes:
        k = (a["data"].year, a["data"].month)
        por_mes[k] = por_mes.get(k, 0) + a["valor"]
    ini = min(por_mes)
    atual = melhor = 0
    for k in _meses_entre(dt.date(ini[0], ini[1], 1), hoje):
        if por_mes.get(k, 0) >= meta_mes - 0.005:
            atual += 1
            melhor = max(melhor, atual)
        elif k != (hoje.year, hoje.month):  # mês em andamento não quebra a sequência
            atual = 0
    return melhor


def _meses_no_azul(e, hoje):
    por_mes = {}
    for g in e.gastos:
        k = (g["data"].year, g["data"].month)
        por_mes[k] = por_mes.get(k, 0) + g["valor"]
    limite = e.salario - e.salario * e.pct_guardar
    return sum(
        1 for k, v in por_mes.items()
        if k < (hoje.year, hoje.month) and e.salario > 0 and v <= limite
    )


def conquistas(e, hoje=None):
    hoje = hoje or dt.date.today()
    r = resumo(e, hoje.year, hoje.month, hoje)
    lista = []

    def add(grupo, nome, desc, atual, alvo):
        lista.append(dict(
            id=f"{grupo}:{alvo}", grupo=grupo, nome=nome, desc=desc,
            atual=atual, alvo=alvo, ok=atual >= alvo,
            prog=min(1.0, atual / alvo) if alvo else 0.0,
        ))

    for alvo, nome in [
        (100, "Primeiro passo"), (500, "Colchão começando"), (1000, "Mil guardados"),
        (5000, "Reserva ganhando forma"), (10000, "Cinco dígitos"),
        (25000, "Quase um carro"), (50000, "Meio caminho dos 100 mil"),
        (100000, "Seis dígitos"),
    ]:
        add("Guardar", nome, f"Guardar {brl(alvo)} no total", r["total_guardado"], alvo)

    for alvo, nome in [
        (1, "Dinheiro trabalhando"), (10, "Primeiro lanche pago pelo investimento"),
        (50, "Rendimento que paga uma conta"), (100, "Cem de renda passiva"),
        (500, "Meio salário mínimo de juros"), (1000, "Mil em rendimentos"),
        (5000, "Os juros viraram salário"),
    ]:
        add("Render", nome, f"Acumular {brl(alvo)} de rendimento", r["rendimento"], alvo)

    for alvo, nome in [(0.25, "Um quarto da meta"), (0.5, "Metade da meta"),
                       (0.75, "Reta final"), (1.0, "Meta batida!")]:
        add("Meta", nome, f"Chegar a {alvo:.0%} da meta '{e.meta_nome}'",
            r["pct_meta"], alvo)

    seq = _sequencia_guardando(e, hoje)
    for alvo, nome in [(3, "Hábito criado"), (6, "Meio ano de disciplina"),
                       (12, "Um ano guardando")]:
        add("Hábito", nome, f"Guardar a meta mensal por {alvo} meses seguidos", seq, alvo)

    azul = _meses_no_azul(e, hoje)
    for alvo, nome in [(1, "Mês no azul"), (3, "Trimestre no controle")]:
        add("Hábito", nome,
            f"Fechar {alvo} mês(es) gastando menos do que sobra depois de guardar", azul, alvo)
    return lista


# ----------------------------------------------------------- investimentos
def taxa_bruta(e, inv):
    v, ix = inv["valor"], inv["indexador"]
    if ix == "%CDI":
        return e.cdi * v / 100
    if ix == "%Selic":
        return e.selic * v / 100
    if ix == "IPCA+":
        return (1 + e.ipca) * (1 + v / 100) - 1
    return v / 100


def ranking_investimentos(e):
    out = []
    for inv in e.investimentos:
        b = taxa_bruta(e, inv)
        out.append(dict(inv, bruto=b, liquido=b * (1 - inv["ir"])))
    return sorted(out, key=lambda x: -x["liquido"])


def simular(e, inv, aporte_mensal, meses, inicial=0.0):
    i = (1 + taxa_bruta(e, inv)) ** (1 / 12) - 1
    saldo, investido = inicial, inicial
    for _ in range(int(meses)):
        saldo = saldo * (1 + i) + aporte_mensal
        investido += aporte_mensal
    ganho = saldo - investido
    liquido = saldo - max(0.0, ganho) * inv["ir"]
    return dict(investido=investido, bruto_final=saldo, liquido_final=liquido,
                ganho_liquido=liquido - investido)


def dicas(e, hoje=None):
    hoje = hoje or dt.date.today()
    r = resumo(e, hoje.year, hoje.month, hoje)
    rk = ranking_investimentos(e)
    media = media_gastos(e)
    base = media if media > 0 else max(0.0, e.salario * (1 - e.pct_guardar))
    out = [
        f"Cenário usado: Selic {e.selic:.2%} a.a., CDI {e.cdi:.2%} a.a., IPCA 12m {e.ipca:.2%}. "
        "Com juros neste nível, a renda fixa paga bem e é o ponto de partida mais simples."
    ]
    if rk:
        top = rk[0]
        out.append(
            f"Maior rendimento líquido estimado na sua lista: {top['nome']} "
            f"(~{top['liquido']:.2%} a.a. líquido). Maior rendimento nem sempre é o melhor: "
            "olhe também liquidez (quando você consegue resgatar) e risco."
        )
    if base > 0:
        faixa = f"{brl(3 * base)} a {brl(6 * base)}"
        if r["patrimonio"] < 3 * base:
            out.append(
                f"1º passo, reserva de emergência (3 a 6 meses de gastos ≈ {faixa}): você tem "
                f"{brl(r['patrimonio'])}. Para essa reserva priorize liquidez diária: "
                "Tesouro Selic ou CDB de liquidez diária com 100% do CDI ou mais."
            )
        elif r["patrimonio"] < 6 * base:
            out.append(
                f"Sua reserva está a caminho (alvo ≈ {faixa}). Complete com liquidez diária e, "
                "depois disso, o que passar do alvo pode ir para prazos maiores (LCI/LCA, CDB de prazo fixo)."
            )
        else:
            out.append(
                f"Sua reserva de emergência já cobre 6 meses (≈ {faixa}). O que passar disso pode ir "
                "para prazos maiores: LCI/LCA, CDB de prazo fixo ou Tesouro IPCA+ (metas de 3+ anos)."
            )
    out.append(
        "Dívida cara (cartão rotativo, cheque especial) custa muito mais do que qualquer "
        "investimento rende: quitar vem antes de investir."
    )
    out.append(
        "A poupança rende bem menos do que as opções acima quando a Selic está alta. "
        "Ações e FIIs não entram neste ranking: oscilam e só fazem sentido para prazos longos (5+ anos), "
        "depois da reserva."
    )
    out.append(
        "Aviso: informação educativa, não recomendação personalizada de investimento. "
        "As taxas são de referência (pesquisa de 25/10/2026) e variam por banco e corretora: "
        "confira no Tesouro Direto e na sua instituição, e atualize a aba Investimentos."
    )
    return out


if __name__ == "__main__":
    if not os.path.exists(ARQUIVO_PADRAO):
        salvar(Estado())
        print("Planilha criada:", ARQUIVO_PADRAO)
