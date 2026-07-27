#!/usr/bin/env python3
"""Gera a apresentação semanal RSN/SSMA a partir da planilha de ocorrências
e da apresentação da semana anterior, sem alterar o layout do template.

Uso:
    python3 gerar_apresentacao.py \
        --planilha SSMA_2026_Dados_e_Ocorrencias_Unificados_v01.xlsx \
        --pptx-anterior Pre_RSN_Atualizada_v9_1.pptx \
        --saida Pre_RSN_Atualizada_v10.pptx

Escopo: slides 1, 3, 4 e 5 (dados de ocorrências das abas "Tirolez" e
"Levitare|Regina"). O slide 2 (metodologia) é copiado sem alteração. O
slide 6 (ETE/meio ambiente) não tem fonte de dados nesta planilha e é
removido da apresentação gerada.
"""
import argparse
import copy
import datetime
import re
from collections import Counter

import openpyxl
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def qn(tag):
    return f"{{{A_NS}}}{tag}"


# --------------------------------------------------------------------------
# Normalização de dados
# --------------------------------------------------------------------------

UNIT_CANON = {
    "arapua": "Arapuá",
    "tiros": "Tiros",
    "monte aprazivel": "Monte Aprazível",
    "caxambu do sul": "Caxambu do Sul",
    "caxambu": "Caxambu do Sul",
    "lins": "Lins",
    "merchandising": "Merchandising",
    "merch": "Merchandising",
    "cd": "CD",
    "cd sp": "CD",
    "vendas": "Vendas",
    "matriz": "Matriz ADM",
    "matriz adm": "Matriz ADM",
    "adm": "Matriz ADM",
    "levitare": "Levitare",
    "regina": "Regina",
}

# rótulo de cada unidade no quadro "EVOLUÇÃO POR UNIDADE" (slide 3)
UNIT_LABEL_SLIDE3 = {
    "Arapuá": "Arapuá",
    "Tiros": "Tiros",
    "Monte Aprazível": "Monte Apr.",
    "Caxambu do Sul": "Caxambu",
    "Lins": "Lins",
    "Merchandising": "Merch.",
    "CD": "CD SP",
    "Vendas": "Vendas",
    "Matriz ADM": "Matriz ADM",
    "Levitare": "Levitare",
    "Regina": "Regina",
}

FABRICA_UNITS = {"Arapuá", "Tiros", "Monte Aprazível", "Lins", "Caxambu do Sul"}
PYRAMID_COLS = ["FÁBRICA", "LEVITARE", "REGINA", "CD", "MERCH", "VENDAS", "ADM"]
COL_UNIT_MAP = {
    "FÁBRICA": FABRICA_UNITS,
    "LEVITARE": {"Levitare"},
    "REGINA": {"Regina"},
    "CD": {"CD"},
    "MERCH": {"Merchandising"},
    "VENDAS": {"Vendas"},
    "ADM": {"Matriz ADM"},
}

SEVERITY_ORDER = ["Irreversível", "ACA", "ASA", "Incidente", "Desvio"]
FAROL_COLOR = {
    "Desvio": "00B050",
    "Incidente": "FFFF00",
    "ASA": "ED7D31",
    "ACA": "C00000",
    "Irreversível": "000000",
}
LEVEL_NAMES = {1: "Patológico", 2: "Reativo", 3: "Calculativo", 4: "Proativo", 5: "Generativo"}

# Calibração da posição vertical (EMU) dos marcadores da régua visual (slide 3),
# lida diretamente das duas posições conhecidas no template original (Shape 215/217
# = linhas indicadoras, Shape 219/221 = rótulos), para nível 1,0 e 1,5. A escala do
# quadro (lv0..lv4) é linear em EMU por nível, então interpolamos/extrapolamos
# linearmente essas duas referências para qualquer nível entre 1 e 5.
_LADDER_LINE_REF = ((1.0, 3673442), (1.5, 3404000))
_LADDER_LABEL_REF = ((1.0, 3606457), (1.5, 3239456))


def _ladder_top(v, ref):
    (v0, t0), (v1, t1) = ref
    return round(t0 + (t1 - t0) * (v - v0) / (v1 - v0))


def ladder_line_top(v):
    return _ladder_top(v, _LADDER_LINE_REF)


def ladder_label_top(v):
    return _ladder_top(v, _LADDER_LABEL_REF)


def avoid_label_overlap(top_a, top_b, min_gap=340000):
    """Os dois rótulos ('Sem. passada'/'Esta semana') têm ~320000 EMU de
    altura; quando os níveis ficam próximos (comum, já que a variação
    semanal é de 0,2 a 0,5), a posição calibrada da escala os faria
    sobrepor. Afasta os dois simetricamente a partir do meio, preservando
    qual dos dois fica visualmente mais acima."""
    if abs(top_b - top_a) >= min_gap:
        return top_a, top_b
    mid = (top_a + top_b) / 2
    if top_a <= top_b:
        return round(mid - min_gap / 2), round(mid + min_gap / 2)
    return round(mid + min_gap / 2), round(mid - min_gap / 2)
MESES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
            "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]


def canon_unit(raw):
    key = (raw or "").strip().lower()
    return UNIT_CANON.get(key, (raw or "").strip())


def norm_classif(c):
    c = (c or "").strip()
    if c.startswith("Irrevers"):
        return "Irreversível"
    if c.startswith("Incidente"):
        return "Incidente"
    return c


def norm_tipo(t):
    t = (t or "").strip()
    if t.startswith("Terceiro"):
        return "Terceiro"
    return t


def fmt_nivel(v):
    return f"{v:.1f}".replace(".", ",")


# --------------------------------------------------------------------------
# Carga da planilha
# --------------------------------------------------------------------------

def load_occurrences(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows = []
    for sheet, bloco in [("Tirolez", "Tirolez"), ("Levitare|Regina", "Levitare|Regina")]:
        ws = wb[sheet]
        for r in range(5, ws.max_row + 1):
            data = ws.cell(r, 2).value
            if not isinstance(data, datetime.datetime):
                continue
            unidade_raw = (ws.cell(r, 3).value or "").strip()
            rows.append(dict(
                data=data,
                unidade=canon_unit(unidade_raw),
                tipo=norm_tipo(ws.cell(r, 4).value),
                classif=norm_classif(ws.cell(r, 5).value),
                desc=(ws.cell(r, 6).value or "").strip(),
                bloco=bloco,
            ))
    return rows


def in_range(row, start, end):
    return start <= row["data"] <= end


def previous_friday_or_same(d):
    """Retrocede `d` até a sexta-feira mais recente (ou mantém, se já for
    sexta). Usado para achar o fechamento normal de semana mesmo quando a
    data mostrada foi estendida pela exceção de ACA (veja classify_week)."""
    days_since_friday = (d.weekday() - 4) % 7  # weekday(): seg=0 ... sex=4
    return d - datetime.timedelta(days=days_since_friday)


def classify_week(row, cur_start, cur_end):
    """Classifica uma ocorrência como da semana 'anterior', 'atual', ou
    nenhuma (None). Semanas normais são sábado a sexta (7 dias), sempre
    contíguas: a semana anterior é sempre cur_start-7 .. cur_start-1.

    Exceção: um ACA que acontece entre sábado e segunda da semana atual
    (os 3 primeiros dias) é reportado de imediato no relatório que está
    sendo fechado agora — conta como 'anterior', não 'atual' — para não
    esperar o ciclo inteiro fechar antes de escalar um acidente grave."""
    prev_start = cur_start - datetime.timedelta(days=7)
    prev_end = cur_start - datetime.timedelta(days=1)
    exception_end = cur_start + datetime.timedelta(days=2)  # sábado + 2 = segunda
    if row["classif"] == "ACA" and cur_start <= row["data"] <= exception_end:
        return "anterior"
    if prev_start <= row["data"] <= prev_end:
        return "anterior"
    if cur_start <= row["data"] <= cur_end:
        return "atual"
    return None


# --------------------------------------------------------------------------
# Extrair intervalo da semana anterior a partir do pptx anterior
# --------------------------------------------------------------------------

def extract_prev_range(prs):
    slide4 = prs.slides[3]
    for shape in slide4.shapes:
        if shape.has_text_frame and "SEMANA" in shape.text_frame.text.upper():
            m = re.search(r"(\d{2})/(\d{2})\s*[–-]\s*(\d{2})/(\d{2})/(\d{4})", shape.text_frame.text)
            if m:
                d1, m1, d2, m2, y = map(int, m.groups())
                return (datetime.datetime(y, m1, d1), datetime.datetime(y, m2, d2))
    raise RuntimeError("Não encontrei o intervalo de datas da semana anterior no slide 4.")


def extract_prev_level(prs):
    slide3 = prs.slides[2]

    def walk(shapes):
        for s in shapes:
            if s.has_text_frame:
                yield s
            if s.shape_type == 6:
                yield from walk(s.shapes)

    for s in walk(slide3.shapes):
        m = re.match(r"Esta semana ([\d,]+)", s.text_frame.text)
        if m:
            return float(m.group(1).replace(",", "."))
    raise RuntimeError("Não encontrei o nível 'Esta semana' da régua no slide 3 anterior.")


# --------------------------------------------------------------------------
# Régua de maturidade — sugestão automática (ver README: revisar antes de publicar)
# --------------------------------------------------------------------------

def suggest_regua(prev_level, week_rows):
    """Aplica as regras do slide 2 (Eixo 2 — modulador semanal) ao nível
    da semana anterior. Só um gatilho vale por semana (o mais severo
    presente), igual ao comportamento observado no template original
    (3 ACA + 1 ASA na mesma semana só produziram -0,5, não a soma).

    IPS/IPA e "quase-acidente relatado" não são medidos por esta
    planilha. Como proxy, uma semana sem nenhum ACA/ASA é tratada como
    evidência de proatividade sustentada e sobe o nível (+0,3, dentro do
    teto de +0,5/ciclo) — em vez de ficar estagnada só porque não há
    dado de IPS/IPA. Acidente irreversível sempre reseta para 1,0."""
    aca_n = sum(1 for r in week_rows if r["classif"] == "ACA")
    asa_n = sum(1 for r in week_rows if r["classif"] == "ASA")
    irrev_n = sum(1 for r in week_rows if r["classif"] == "Irreversível")

    if irrev_n > 0:
        new_level = 1.0
        motivos = ["acidente irreversível — reset do nível para 1,0"]
    elif aca_n > 0:
        motivos = [f"{aca_n} ACA (teto -0,5)"]
        new_level = max(1.0, min(5.0, round(prev_level - 0.5, 1)))
    elif asa_n > 0:
        motivos = [f"{asa_n} ASA (-0,2)"]
        new_level = max(1.0, min(5.0, round(prev_level - 0.2, 1)))
    else:
        motivos = ["semana sem ACA/ASA — proatividade sustentada (proxy de IPS/IPA ≥90%, +0,3)"]
        new_level = max(1.0, min(5.0, round(prev_level + 0.3, 1)))

    return dict(
        prev_level=prev_level, new_level=new_level, delta=round(new_level - prev_level, 2),
        motivos=motivos, aca_n=aca_n, asa_n=asa_n, irrev_n=irrev_n,
        nota="SUGESTÃO AUTOMÁTICA — IPS/IPA real e quase-acidentes relatados não estão na planilha; a subida usa uma semana limpa de ACA/ASA como proxy. Revisar com o time de SSMA antes de publicar.",
    )


# --------------------------------------------------------------------------
# Helpers de edição de texto preservando formatação
# --------------------------------------------------------------------------

def set_run_text(shape, new_text):
    """Substitui o texto do primeiro run, preservando a formatação, e
    remove runs adicionais (o texto alvo é sempre uma única linha curta)."""
    tf = shape.text_frame
    para = tf.paragraphs[0]
    if not para.runs:
        para.text = new_text
    else:
        para.runs[0].text = new_text
        for extra in para.runs[1:]:
            extra._r.getparent().remove(extra._r)
    # remove parágrafos extras porventura existentes (evita texto antigo "fantasma")
    for extra_p in tf.paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)


def set_multiline_text(shape, lines):
    """Como set_run_text, mas para shapes com um parágrafo por linha
    (ex.: cartões ACA com título + descrição), preservando a formatação
    de cada parágrafo."""
    tf = shape.text_frame
    paras = list(tf.paragraphs)
    for i, line in enumerate(lines):
        if i < len(paras):
            p = paras[i]
        else:
            p_el = copy.deepcopy(paras[-1]._p)
            paras[-1]._p.addnext(p_el)
            paras = list(tf.paragraphs)
            p = paras[i]
        if p.runs:
            p.runs[0].text = line
            for extra in p.runs[1:]:
                extra._r.getparent().remove(extra._r)
        else:
            p.text = line
    paras = list(tf.paragraphs)
    for p in paras[len(lines):]:
        p._p.getparent().remove(p._p)


def find_by_name(shapes, name):
    for s in shapes:
        if s.name == name:
            return s
        if s.shape_type == 6:
            r = find_by_name(s.shapes, name)
            if r:
                return r
    return None


def walk_all(shapes):
    for s in shapes:
        yield s
        if s.shape_type == 6:
            yield from walk_all(s.shapes)


def find_by_text(shapes, text):
    for s in walk_all(shapes):
        if s.has_text_frame and s.text_frame.text.strip() == text:
            return s
    return None


def delete_shape(shape):
    shape._element.getparent().remove(shape._element)


def delete_slide(prs, index):
    """Remove o slide `index` (0-based) da apresentação, junto com o
    relacionamento correspondente em presentation.xml.rels."""
    sldIdLst = prs.slides._sldIdLst
    slide_id_elements = list(sldIdLst)
    r_id = slide_id_elements[index].attrib[
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    ]
    prs.part.drop_rel(r_id)
    sldIdLst.remove(slide_id_elements[index])


def set_cell_text_keep_format(cell, new_text):
    tf = cell.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = new_text
        for extra in para.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        para.text = new_text


def set_cell_farol(cell, classif):
    para = cell.text_frame.paragraphs[0]
    run = para.runs[0]
    color = FAROL_COLOR.get(classif, "808080")
    srgb = run._r.find(qn("rPr")).find(qn("solidFill")).find(qn("srgbClr"))
    srgb.set("val", color)


# --------------------------------------------------------------------------
# Slide 1 — capa
# --------------------------------------------------------------------------

def edit_slide1(prs, start, end):
    slide = prs.slides[0]
    for shape in slide.shapes:
        if shape.has_text_frame and re.match(r"\d{2}/\d{2} a \d{2}/\d{2}", shape.text_frame.text):
            set_run_text(shape, f"{start.strftime('%d/%m')} a {end.strftime('%d/%m')}")


# --------------------------------------------------------------------------
# Slide 3 — termômetro de cultura
# --------------------------------------------------------------------------

def build_unit_matrix(rows):
    units = sorted({r["unidade"] for r in rows if r["semana"] in ("anterior", "atual")})
    data = []
    for u in units:
        prev_rows = [r for r in rows if r["unidade"] == u and r["semana"] == "anterior"]
        cur_rows = [r for r in rows if r["unidade"] == u and r["semana"] == "atual"]

        def counts(rs):
            d = sum(1 for r in rs if r["classif"] == "Desvio")
            i = sum(1 for r in rs if r["classif"] == "Incidente")
            a = sum(1 for r in rs if r["classif"] == "ACA")
            return d, i, a, len(rs)

        pd_, pi_, pa_, pt_ = counts(prev_rows)
        cd_, ci_, ca_, ct_ = counts(cur_rows)
        data.append(dict(unidade=u, prev=(pd_, pi_, pa_, pt_), cur=(cd_, ci_, ca_, ct_)))
    data.sort(key=lambda x: (-x["cur"][3], -x["prev"][3], x["unidade"]))
    return data


def arrow_cell_text(prev, cur):
    if prev == 0 and cur == 0:
        return "–"
    if cur > prev:
        return f"{prev}→{cur} ▲"
    if cur < prev:
        return f"{prev}→{cur} ▼"
    return f"{prev}→{cur}"


def edit_slide3(prs, rows, regua):
    slide = prs.slides[2]
    all_shapes = list(slide.shapes)

    # --- régua ---
    lvl_prev = regua["prev_level"]
    lvl_new = regua["new_level"]
    set_run_text(find_by_name(all_shapes, "Shape 219"), f"Sem. passada {fmt_nivel(lvl_prev)}")
    set_run_text(find_by_name(all_shapes, "Shape 221"), f"Esta semana {fmt_nivel(lvl_new)}")

    # reposiciona os marcadores na escala visual da régua (1 a 5) — o texto por si só
    # não move o indicador; sem isso, os marcadores ficam presos na posição da semana anterior.
    find_by_name(all_shapes, "Shape 215").top = Emu(ladder_line_top(lvl_prev))
    find_by_name(all_shapes, "Shape 217").top = Emu(ladder_line_top(lvl_new))
    label_prev_top, label_new_top = avoid_label_overlap(ladder_label_top(lvl_prev), ladder_label_top(lvl_new))
    find_by_name(all_shapes, "Shape 219").top = Emu(label_prev_top)
    find_by_name(all_shapes, "Shape 221").top = Emu(label_new_top)
    if regua["delta"] > 0:
        arrow = "▲"
    elif regua["delta"] < 0:
        arrow = "▼"
    else:
        arrow = "="
    delta_txt = fmt_nivel(regua["delta"]) if regua["delta"] != 0 else "0,0"
    nome_prev = LEVEL_NAMES[round(lvl_prev)]
    nome_new = LEVEL_NAMES[round(lvl_new)]
    if regua["delta"] == 0:
        arrow_text = f"= sem variação · {nome_new}"
    elif nome_prev == nome_new:
        arrow_text = f"{arrow} {delta_txt} · dentro de {nome_new}"
    else:
        arrow_text = f"{arrow} {delta_txt} · {nome_prev} → {nome_new}"
    set_run_text(find_by_name(all_shapes, "TextBox 223"), arrow_text)

    # --- ocorrências por tipo (grupo 'Agrupar 5') ---
    cur_rows = [r for r in rows if r["semana"] == "atual"]
    prev_rows = [r for r in rows if r["semana"] == "anterior"]

    def count(rs, classif):
        return sum(1 for r in rs if r["classif"] == classif)

    metric_map = {
        "Desvios": "Desvio",
        "Incidentes": "Incidente",
        "Acid. s/ afast.": "ASA",
        "Acid. c/ afast.": "ACA",
        "Irreversível": "Irreversível",
    }
    grupo = find_by_name(slide.shapes, "Agrupar 5")
    children = list(grupo.shapes)
    for idx, s in enumerate(children):
        if s.has_text_frame and s.text_frame.text.strip() in metric_map:
            classif = metric_map[s.text_frame.text.strip()]
            valores = [c for c in children[idx + 1:] if c.has_text_frame and c.text_frame.text.strip()]
            prev_shape, cur_shape = valores[0], valores[1]
            set_run_text(prev_shape, str(count(prev_rows, classif)))
            set_run_text(cur_shape, str(count(cur_rows, classif)))

    # --- headline ---
    aca_n, asa_n, irrev_n = regua["aca_n"], regua["asa_n"], regua["irrev_n"]
    if aca_n + asa_n + irrev_n == 0:
        destaque = "nenhum ACA/ASA na semana"
    else:
        tipos_graves = Counter(r["tipo"] for r in cur_rows if r["classif"] in ("ACA", "ASA", "Irreversível"))
        top_tipo = tipos_graves.most_common(1)[0][0].lower()
        destaque = f"{top_tipo} lidera as ocorrências graves"
    resumo_acidentes = destaque if aca_n + asa_n == 0 else (
        f"{aca_n} ACA{'s' if aca_n != 1 else ''} e {asa_n} ASA · {destaque}"
    )
    if regua["delta"] == 0:
        headline = f"Cultura estável em {nome_new} ({fmt_nivel(lvl_new)}); {resumo_acidentes}"
    elif nome_prev == nome_new:
        direcao = "sobe" if regua["delta"] > 0 else "recua"
        headline = (
            f"Cultura {direcao} dentro do nível {nome_new} "
            f"({fmt_nivel(lvl_prev)}→{fmt_nivel(lvl_new)}); {resumo_acidentes}"
        )
    else:
        direcao = "sobe" if regua["delta"] > 0 else "recua"
        headline = (
            f"Cultura {direcao} de {nome_prev} para {nome_new} "
            f"({fmt_nivel(lvl_prev)}→{fmt_nivel(lvl_new)}); {resumo_acidentes}"
        )
    set_run_text(find_by_name(all_shapes, "TextBox 3"), headline)

    # --- quadro EVOLUÇÃO POR UNIDADE ---
    matrix = build_unit_matrix(rows)
    n_new = len(matrix)
    n_old = 8
    row_pitch = Emu(152000)
    if n_new > n_old:
        raise RuntimeError(
            f"{n_new} unidades ativas excedem as {n_old} vagas do quadro 'EVOLUÇÃO POR UNIDADE'. "
            "Layout precisa de ajuste manual (adicionar linhas) antes de rodar o agente."
        )

    for i in range(n_new, n_old):
        for name in [f"u{i}", f"c{i}-Desvios", f"c{i}-Incidentes", f"c{i}-ACA", f"c{i}-Total"]:
            sh = find_by_name(slide.shapes, name)
            if sh is not None:
                delete_shape(sh)

    for i, row in enumerate(matrix):
        set_run_text(find_by_name(slide.shapes, f"u{i}"), UNIT_LABEL_SLIDE3.get(row["unidade"], row["unidade"]))
        pd_, pi_, pa_, pt_ = row["prev"]
        cd_, ci_, ca_, ct_ = row["cur"]
        set_run_text(find_by_name(slide.shapes, f"c{i}-Desvios"), arrow_cell_text(pd_, cd_))
        set_run_text(find_by_name(slide.shapes, f"c{i}-Incidentes"), arrow_cell_text(pi_, ci_))
        set_run_text(find_by_name(slide.shapes, f"c{i}-ACA"), arrow_cell_text(pa_, ca_))
        set_run_text(find_by_name(slide.shapes, f"c{i}-Total"), arrow_cell_text(pt_, ct_))

    shrink = (n_old - n_new) * row_pitch
    if shrink:
        for name in ["g-rule", "g-band", "g-label", "g-Desvios", "g-Incidentes", "g-ACA", "g-Total"]:
            sh = find_by_name(slide.shapes, name)
            sh.top = Emu(sh.top - shrink)
        card = find_by_name(slide.shapes, "Matriz card")
        card.height = Emu(card.height - shrink)

    # GERAL = soma das linhas ativas
    g_d = sum(r["cur"][0] for r in matrix)
    g_i = sum(r["cur"][1] for r in matrix)
    g_a = sum(r["cur"][2] for r in matrix)
    g_t = sum(r["cur"][3] for r in matrix)
    pg_d = sum(r["prev"][0] for r in matrix)
    pg_i = sum(r["prev"][1] for r in matrix)
    pg_a = sum(r["prev"][2] for r in matrix)
    pg_t = sum(r["prev"][3] for r in matrix)
    set_run_text(find_by_name(slide.shapes, "g-Desvios"), arrow_cell_text(pg_d, g_d))
    set_run_text(find_by_name(slide.shapes, "g-Incidentes"), arrow_cell_text(pg_i, g_i))
    set_run_text(find_by_name(slide.shapes, "g-ACA"), arrow_cell_text(pg_a, g_a))
    set_run_text(find_by_name(slide.shapes, "g-Total"), arrow_cell_text(pg_t, g_t))

    # --- resumo (TOTAIS DA SEMANA) ---
    acidentes_prev = count(prev_rows, "ACA") + count(prev_rows, "ASA")
    acidentes_cur = aca_n + asa_n
    if g_t == pg_t:
        delta_tot_txt = "± 0"
    elif g_t > pg_t:
        delta_tot_txt = f"+{g_t - pg_t}"
    else:
        delta_tot_txt = f"−{pg_t - g_t}"
    set_run_text(
        find_by_name(slide.shapes, "r-tot"),
        f"Ocorrências {pg_t} → {g_t}  {delta_tot_txt}",
    )
    acc_arrow = "▲" if acidentes_cur > acidentes_prev else ("▼" if acidentes_cur < acidentes_prev else "")
    aca_arrow = "▲" if aca_n > count(prev_rows, "ACA") else ("▼" if aca_n < count(prev_rows, "ACA") else "")
    asa_arrow = "▲" if asa_n > count(prev_rows, "ASA") else ("▼" if asa_n < count(prev_rows, "ASA") else "")
    set_run_text(
        find_by_name(slide.shapes, "r-acc"),
        f"Acidentes {acidentes_prev} → {acidentes_cur} {acc_arrow}  ·  ACA {count(prev_rows,'ACA')} → {aca_n} {aca_arrow}  ·  ASA {count(prev_rows,'ASA')} → {asa_n} {asa_arrow}",
    )
    asa_destaques = [r for r in cur_rows if r["classif"] == "ASA"]
    nota_extra = ""
    if asa_destaques:
        exemplo = asa_destaques[0]
        nota_extra = f" · +{len(asa_destaques)} ASA ({UNIT_LABEL_SLIDE3.get(exemplo['unidade'], exemplo['unidade'])}/{exemplo['tipo'].lower()})."
    set_run_text(
        find_by_name(slide.shapes, "r-note"),
        f"Irreversível: {irrev_n} no período{nota_extra if nota_extra else '.'}",
    )


# --------------------------------------------------------------------------
# Slide 4 — pirâmide de segurança YTD
# --------------------------------------------------------------------------

def pyramid_count(rows, classif_target, col, until):
    sel = [r for r in rows if r["classif"] == classif_target and r["data"] <= until]
    if col == "TRAJETO":
        return sum(1 for r in sel if r["tipo"] == "Trajeto")
    if col == "TERCEIRO":
        return sum(1 for r in sel if r["tipo"] == "Terceiro")
    units = COL_UNIT_MAP[col]
    return sum(1 for r in sel if r["unidade"] in units and r["tipo"] not in ("Trajeto", "Terceiro"))


def edit_slide4(prs, rows, cur_range, regua):
    slide = prs.slides[3]
    all_shapes = list(slide.shapes)
    start, end = cur_range
    iso_week = end.isocalendar()[1]
    mes = MESES_PT[end.month - 1]

    for shape in all_shapes:
        if shape.has_text_frame and "SEMANA" in shape.text_frame.text.upper() and "VAMOS" in shape.text_frame.text.upper():
            set_run_text(shape, f"VAMOS FALAR DE SEGURANÇA?     SEMANA {start.strftime('%d/%m')} – {end.strftime('%d/%m/%Y')} · YTD atualizado")

    cur_rows = [r for r in rows if r["semana"] == "atual"]
    aca_n, asa_n, irrev_n = regua["aca_n"], regua["asa_n"], regua["irrev_n"]

    # tabela pirâmide
    tbl = None
    for shape in all_shapes:
        if shape.has_table:
            tbl = shape.table
    ytd = {}
    for ri, classif in enumerate(SEVERITY_ORDER, start=1):
        vals = [pyramid_count(rows, classif, col, end) for col in PYRAMID_COLS]
        trajeto = pyramid_count(rows, classif, "TRAJETO", end)
        terceiro = pyramid_count(rows, classif, "TERCEIRO", end)
        s_semana = sum(1 for r in rows if r["classif"] == classif and r["data"].isocalendar()[1] == iso_week and r["data"].isocalendar()[0] == end.isocalendar()[0])
        s_mes = sum(1 for r in rows if r["classif"] == classif and r["data"].year == end.year and r["data"].month == end.month)
        ytd_total = sum(vals) + trajeto + terceiro
        linha = vals + [trajeto, terceiro, s_semana, s_mes, ytd_total]
        ytd[classif] = ytd_total
        for ci, v in enumerate(linha):
            texto = "-" if v == 0 else str(v)
            set_cell_text_keep_format(tbl.rows[ri].cells[ci], texto)
    tbl.rows[0].cells[9].text_frame.paragraphs[0].runs[0].text = f"S {iso_week}"
    tbl.rows[0].cells[10].text_frame.paragraphs[0].runs[0].text = mes

    # KPIs (Irreversível / ACA / ASA / Incidentes / Desvios) — valor é o shape imediatamente anterior ao rótulo
    kpi_labels = {
        "Irreversível": ytd["Irreversível"],
        "Acid. COM afastamento": ytd["ACA"],
        "Acid. SEM afastamento": ytd["ASA"],
        "Incidentes": ytd["Incidente"],
        "Desvios": ytd["Desvio"],
    }
    for idx, shape in enumerate(all_shapes):
        if shape.has_text_frame and shape.text_frame.text.strip() in kpi_labels:
            valor_shape = all_shapes[idx - 1]
            set_run_text(valor_shape, str(kpi_labels[shape.text_frame.text.strip()]))

    # headline / subtítulo
    if aca_n + asa_n + irrev_n == 0:
        irrev_txt = f"{ytd['Irreversível']} irreversível{'is' if ytd['Irreversível'] != 1 else ''}"
        headline = (
            f"Semana sem acidentes com ou sem afastamento; YTD permanece em {ytd['ACA']} ACAs e "
            f"{irrev_txt} — manter vigilância nos pontos já mapeados"
        )
        subtitulo = "OCORRÊNCIAS MAIS GRAVES DA SEMANA · nenhum ACA/ASA/irreversível registrado"
    else:
        partes = []
        if aca_n:
            partes.append(f"{aca_n} ACA{'s' if aca_n != 1 else ''}")
        if asa_n:
            partes.append(f"{asa_n} ASA")
        if irrev_n:
            partes.append(f"{irrev_n} irreversível{'is' if irrev_n != 1 else ''}")
        tipos_graves = Counter(r["tipo"] for r in cur_rows if r["classif"] in ("ACA", "ASA", "Irreversível"))
        causa = tipos_graves.most_common(1)[0][0].lower()
        headline = f"Semana com {' e '.join(partes)} eleva{'m' if len(partes)>1 else ''} o YTD para {ytd['ACA']} ACAs; {causa} exige atenção imediata"
        graves = [r for r in cur_rows if r["classif"] in ("ACA", "ASA", "Irreversível")]
        datas = sorted(r["data"] for r in graves)
        subtitulo = (
            f"OCORRÊNCIAS MAIS GRAVES DA SEMANA · {aca_n} ACA(s) registrados entre "
            f"{datas[0].strftime('%d/%m')} e {datas[-1].strftime('%d/%m/%Y')}"
        )
    set_run_text(find_by_name(all_shapes, "TextBox 2"), headline)
    set_run_text(find_by_name(all_shapes, "TextBox 22"), subtitulo)

    # cards de ocorrências mais graves (ACA / Irreversível)
    graves = sorted(
        [r for r in cur_rows if r["classif"] in ("ACA", "Irreversível")],
        key=lambda r: r["data"],
    )
    n_slots = 3
    if len(graves) > n_slots:
        raise RuntimeError(
            f"{len(graves)} ACA/irreversíveis na semana excedem os {n_slots} cartões disponíveis no slide 4. "
            "Ajuste manual de layout necessário."
        )
    if not graves:
        card_bar = find_by_name(all_shapes, "Rectangle_ACA_bar_1")
        card_tx = find_by_name(all_shapes, "TextBox_ACA_1")
        card_bar.fill.solid()
        card_bar.fill.fore_color.rgb = RGBColor.from_string("00B050")
        set_multiline_text(card_tx, ["● Sem ACA/irreversível esta semana", "Nenhum acidente com afastamento ou evento irreversível registrado no período. Manter o padrão."])
        for i in (2, 3):
            for prefix in ("Rectangle_ACA_bar_", "Rectangle_ACA_bg_", "TextBox_ACA_"):
                sh = find_by_name(all_shapes, f"{prefix}{i}")
                if sh is not None:
                    delete_shape(sh)
    else:
        for i in range(len(graves), n_slots):
            for prefix in ("Rectangle_ACA_bar_", "Rectangle_ACA_bg_", "TextBox_ACA_"):
                sh = find_by_name(all_shapes, f"{prefix}{i + 1}")
                if sh is not None:
                    delete_shape(sh)
        for i, r in enumerate(graves, start=1):
            tx = find_by_name(all_shapes, f"TextBox_ACA_{i}")
            label = "ACA" if r["classif"] == "ACA" else "IRREVERSÍVEL"
            set_multiline_text(tx, [
                f"● {label} · {r['data'].strftime('%d/%m/%Y')}  |  {UNIT_LABEL_SLIDE3.get(r['unidade'], r['unidade'])}  |  {r['tipo']}",
                r["desc"],
            ])


# --------------------------------------------------------------------------
# Slide 5 — log detalhado da semana
# --------------------------------------------------------------------------

def resize_table_block(tbl, first_data_row_idx, n_old, n_new):
    """Clona ou remove linhas <a:tr> a partir de first_data_row_idx para
    igualar n_new, preservando a formatação de uma linha de dados existente."""
    tbl_el = tbl._tbl
    trs = tbl_el.findall(qn("tr"))
    template_tr = trs[first_data_row_idx]
    if n_new > n_old:
        anchor = template_tr
        for _ in range(n_new - n_old):
            new_tr = copy.deepcopy(template_tr)
            anchor.addnext(new_tr)
            anchor = new_tr
    elif n_new < n_old:
        for _ in range(n_old - n_new):
            trs = tbl_el.findall(qn("tr"))
            doomed = trs[first_data_row_idx + n_new]
            tbl_el.remove(doomed)


def edit_slide5(prs, rows, cur_range):
    slide = prs.slides[4]
    tbl_shape = next(s for s in slide.shapes if s.has_table)
    tbl = tbl_shape.table

    tirolez_rows = sorted([r for r in rows if r["semana"] == "atual" and r["bloco"] == "Tirolez"], key=lambda r: r["data"])
    levreg_rows = sorted([r for r in rows if r["semana"] == "atual" and r["bloco"] == "Levitare|Regina"], key=lambda r: r["data"])

    # dimensões atuais: header(3 linhas: título+subtítulo+cabeçalho) + N1 + separador(1) + N2 + legenda(1)
    total_rows_before = len(tbl.rows)
    n1_old = None
    for i in range(3, total_rows_before):
        cell0 = tbl.rows[i].cells[0].text
        if cell0 not in ("⬤",):
            n1_old = i - 3
            break
    n2_old = total_rows_before - 1 - 3 - n1_old - 1  # total - legenda - header - N1 - separador

    resize_table_block(tbl, 3, n1_old, len(tirolez_rows))
    sep_idx = 3 + len(tirolez_rows)
    resize_table_block(tbl, sep_idx + 1, n2_old, len(levreg_rows))

    start, end = cur_range
    total = len(tirolez_rows) + len(levreg_rows)
    set_cell_text_keep_format(tbl.rows[0].cells[0], f"OCORRÊNCIAS DA SEMANA · {start.strftime('%d/%m')} – {end.strftime('%d/%m/%Y')}")
    set_cell_text_keep_format(
        tbl.rows[1].cells[0],
        f"Total de ocorrências: {total}   |   Laticínios Tirolez Ltda. (Tirolez · Levitare · Regina)",
    )

    for i, r in enumerate(tirolez_rows):
        row = tbl.rows[3 + i]
        set_cell_text_keep_format(row.cells[1], r["data"].strftime("%d/%m/%Y"))
        set_cell_text_keep_format(row.cells[2], UNIT_LABEL_SLIDE3.get(r["unidade"], r["unidade"]))
        set_cell_text_keep_format(row.cells[3], r["tipo"])
        set_cell_text_keep_format(row.cells[4], r["classif"])
        set_cell_text_keep_format(row.cells[5], r["desc"])
        set_cell_farol(row.cells[0], r["classif"])

    for i, r in enumerate(levreg_rows):
        row = tbl.rows[sep_idx + 1 + i]
        set_cell_text_keep_format(row.cells[1], r["data"].strftime("%d/%m/%Y"))
        set_cell_text_keep_format(row.cells[2], UNIT_LABEL_SLIDE3.get(r["unidade"], r["unidade"]))
        set_cell_text_keep_format(row.cells[3], r["tipo"])
        set_cell_text_keep_format(row.cells[4], r["classif"])
        set_cell_text_keep_format(row.cells[5], r["desc"])
        set_cell_farol(row.cells[0], r["classif"])


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--planilha", required=True)
    ap.add_argument("--pptx-anterior", required=True)
    ap.add_argument("--saida", required=True)
    ap.add_argument("--inicio", help="dd/mm/aaaa — sobrepõe a inferência automática")
    ap.add_argument("--fim", help="dd/mm/aaaa — sobrepõe a inferência automática")
    args = ap.parse_args()

    prs = Presentation(args.pptx_anterior)
    prev_start_shown, prev_end_shown = extract_prev_range(prs)
    if args.inicio and args.fim:
        cur_start = datetime.datetime.strptime(args.inicio, "%d/%m/%Y")
        cur_end = datetime.datetime.strptime(args.fim, "%d/%m/%Y")
    else:
        # semanas normais são sábado a sexta (7 dias); previous_friday_or_same
        # também acerta o caso em que a semana anterior foi mostrada "estendida"
        # por ter puxado um ACA de sábado/domingo/segunda pra dentro dela.
        normal_friday = previous_friday_or_same(prev_end_shown)
        cur_start = normal_friday + datetime.timedelta(days=1)
        cur_end = cur_start + datetime.timedelta(days=6)
    cur_range = (cur_start, cur_end)
    prev_level = extract_prev_level(prs)

    rows = load_occurrences(args.planilha)
    for r in rows:
        r["semana"] = classify_week(r, cur_start, cur_end)
    cur_rows = [r for r in rows if r["semana"] == "atual"]
    regua = suggest_regua(prev_level, cur_rows)

    normal_prev_start = cur_start - datetime.timedelta(days=7)
    normal_prev_end = cur_start - datetime.timedelta(days=1)
    puxados = [r for r in rows if r["classif"] == "ACA" and cur_start <= r["data"] <= cur_start + datetime.timedelta(days=2)]

    print(f"Semana anterior (relatório já publicado): {prev_start_shown:%d/%m/%Y} a {prev_end_shown:%d/%m/%Y}")
    print(f"Semana anterior (janela normal p/ contagem): {normal_prev_start:%d/%m/%Y} a {normal_prev_end:%d/%m/%Y}")
    print(f"Semana atual:    {cur_start:%d/%m/%Y} a {cur_end:%d/%m/%Y}")
    print(f"Ocorrências na semana atual: {len(cur_rows)}")
    if puxados:
        detalhe = "; ".join(f"{r['unidade']} {r['data']:%d/%m}" for r in puxados)
        print(f"ACA(s) puxado(s) para o relatório anterior (não duplicado aqui): {detalhe}")
    print(f"Régua: {fmt_nivel(regua['prev_level'])} -> {fmt_nivel(regua['new_level'])}  ({regua['nota']})")
    if regua["motivos"]:
        print("Motivos:", "; ".join(regua["motivos"]))

    edit_slide1(prs, cur_start, cur_end)
    edit_slide3(prs, rows, regua)
    edit_slide4(prs, rows, cur_range, regua)
    edit_slide5(prs, rows, cur_range)
    delete_slide(prs, 5)  # slide 6 — ETE/meio ambiente: sem fonte de dados, removido
    print("Slide 6 (ETE/meio ambiente) removido — sem fonte de dados nesta planilha.")

    prs.save(args.saida)
    print(f"Gerado: {args.saida}")


if __name__ == "__main__":
    main()
