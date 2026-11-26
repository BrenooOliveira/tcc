"""Conversão de LaTeX para OMML (equações nativas do DOCX).

Estratégia: LaTeX → MathML (pacote opcional latex2mathml) → OMML.
Se a conversão não estiver disponível ou falhar, o chamador deve registrar
pendência e inserir a expressão em fonte monoespaçada, ainda delimitada.
"""

from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


class OmmlError(Exception):
    """Falha explícita na conversão LaTeX → OMML."""


_NARY = set('∑∫∏∐⋃⋂')
_RELATIONS = set('=≠<>≤≥≈±→←∶:')
_OVERLINE = set('\u00af\u203e\u2015\u0305\u0304‾―¯')

_RE_ARG = re.compile(r'\\arg\b')


def append_equation(paragraph, latex: str, display: bool = False):
    """Anexa uma equação OMML ao parágrafo. Levanta OmmlError se falhar."""
    paragraph._p.append(latex_to_omml(latex, display=display))


def latex_to_omml(latex: str, display: bool = False):
    """Devolve o elemento m:oMath ou m:oMathPara."""
    expr = _prepare_latex(latex)
    if not expr:
        raise OmmlError('Expressão LaTeX vazia.')

    try:
        from latex2mathml.converter import convert as latex_to_mathml
    except Exception as exc:
        raise OmmlError(
            'Pacote latex2mathml não está instalado. '
            'Instale com: pip install latex2mathml'
        ) from exc

    try:
        mathml = latex_to_mathml(expr)
        root = etree.fromstring(mathml.encode('utf-8'))
    except Exception as exc:
        raise OmmlError(f'LaTeX não pôde ser convertido para MathML: {exc}') from exc

    omath = _m('oMath')
    for element in _convert(root):
        omath.append(element)
    if len(omath) == 0:
        raise OmmlError('MathML não produziu conteúdo OMML.')

    if not display:
        return omath

    wrapper = _m('oMathPara')
    props = _m('oMathParaPr')
    align = _m('jc')
    align.set(qn('m:val'), 'center')
    props.append(align)
    wrapper.append(props)
    wrapper.append(omath)
    return wrapper


def _prepare_latex(latex: str) -> str:
    expr = (latex or '').strip()
    # latex2mathml deixa \arg literal; \operatorname é reconhecido.
    return _RE_ARG.sub(r'\\operatorname{arg}', expr)


def _m(tag: str):
    return OxmlElement(f'm:{tag}')


def _local(node) -> str:
    return etree.QName(node).localname


def _children(node):
    return [child for child in node if isinstance(child.tag, str)]


def _text(node) -> str:
    return ''.join(node.itertext()).strip()


def _run(text: str, sty: str | None = None):
    run = _m('r')
    if sty:
        props = _m('rPr')
        style = _m('sty')
        style.set(qn('m:val'), sty)
        props.append(style)
        run.append(props)
    token = _m('t')
    token.set(qn('xml:space'), 'preserve')
    token.text = text
    run.append(token)
    return run


def _group(elements):
    holder = _m('e')
    if not elements:
        holder.append(_run(''))
    else:
        for element in elements:
            holder.append(element)
    return holder


def _convert(node) -> list:
    tag = _local(node)
    if tag == 'semantics':
        kids = _children(node)
        return _convert(kids[0]) if kids else []
    if tag in {'math', 'mrow', 'mstyle', 'mpadded', 'menclose'}:
        return _convert_row(_children(node))
    if tag in {'mi', 'mn', 'mo', 'mtext', 'ms'}:
        return [_token(node)]
    if tag == 'mfrac':
        return [_fraction(node)]
    if tag == 'msqrt':
        return [_radical(node, index=None)]
    if tag == 'mroot':
        kids = _children(node)
        index = kids[1] if len(kids) > 1 else None
        base = kids[0] if kids else None
        return [_radical_parts(base, index)]
    if tag in {'msub', 'msup', 'msubsup'}:
        return [_script(node)]
    if tag == 'mover':
        return [_accent(node)]
    if tag == 'munder':
        kids = _children(node)
        if len(kids) >= 2:
            return [_script_from_parts(kids[0], kids[1], None)]
        return _convert_row(kids)
    if tag == 'munderover':
        kids = _children(node)
        base = kids[0] if kids else None
        sub = kids[1] if len(kids) > 1 else None
        sup = kids[2] if len(kids) > 2 else None
        if base is not None and _operator_char(base) in _NARY:
            return [_nary(base, sub, sup, [])]
        return [_script_from_parts(base, sub, sup)]
    if tag == 'mtable':
        return [_matrix(node)]
    if tag == 'mfenced':
        return [_delimiter(node.get('open', '('), node.get('close', ')'), _convert_row(_children(node)))]
    if tag in {'mtd', 'mtr'}:
        return _convert_row(_children(node))
    if _children(node):
        return _convert_row(_children(node))
    if _text(node):
        return [_run(_text(node), 'p')]
    return []


def _convert_row(nodes) -> list:
    output = []
    index = 0
    while index < len(nodes):
        node = nodes[index]
        tag = _local(node)
        if (
            tag == 'mo'
            and _text(node) == '{'
            and index + 1 < len(nodes)
            and _local(nodes[index + 1]) == 'mtable'
        ):
            output.append(_delimiter('{', '', _convert(nodes[index + 1])))
            index += 2
            continue
        if tag in {'msub', 'msup', 'msubsup'} and _operator_char(_children(node)[0] if _children(node) else node) in _NARY:
            operand = []
            cursor = index + 1
            while cursor < len(nodes) and not _is_relation(nodes[cursor]):
                operand.extend(_convert(nodes[cursor]))
                cursor += 1
            output.append(_nary_from_script(node, operand))
            index = cursor
            continue
        output.extend(_convert(node))
        index += 1
    return output


def _token(node):
    tag = _local(node)
    variant = (node.get('mathvariant') or '').lower()
    sty = {'bold': 'b', 'italic': 'i', 'bold-italic': 'bi', 'normal': 'p'}.get(variant)
    if tag in {'mo', 'mn', 'mtext', 'ms'}:
        sty = sty or 'p'
    return _run(_text(node) or '', sty)


def _fraction(node):
    kids = _children(node)
    frac = _m('f')
    props = _m('fPr')
    kind = _m('type')
    kind.set(qn('m:val'), 'bar')
    props.append(kind)
    frac.append(props)
    frac.append(_slot('num', _convert(kids[0]) if kids else []))
    frac.append(_slot('den', _convert(kids[1]) if len(kids) > 1 else []))
    return frac


def _slot(tag: str, elements):
    holder = _m(tag)
    if not elements:
        holder.append(_run(''))
    else:
        for element in elements:
            holder.append(element)
    return holder


def _radical(node, index):
    return _radical_parts(node, index)


def _radical_parts(base, index):
    rad = _m('rad')
    props = _m('radPr')
    if index is None:
        hide = _m('degHide')
        hide.set(qn('m:val'), '1')
        props.append(hide)
    rad.append(props)
    degree = _m('deg')
    if index is not None:
        for element in _convert(index):
            degree.append(element)
    rad.append(degree)
    base_elements = _convert_row(_children(base)) if base is not None and _local(base) == 'msqrt' else (_convert(base) if base is not None else [])
    rad.append(_group(base_elements))
    return rad


def _script(node):
    kids = _children(node)
    tag = _local(node)
    base = kids[0] if kids else None
    sub = kids[1] if tag in {'msub', 'msubsup'} and len(kids) > 1 else None
    if tag == 'msup':
        sup = kids[1] if len(kids) > 1 else None
        sub = None
    elif tag == 'msubsup':
        sup = kids[2] if len(kids) > 2 else None
    else:
        sup = None
    if base is not None and _operator_char(base) in _NARY:
        return _nary(base, sub, sup, [])
    return _script_from_parts(base, sub, sup)


def _script_from_parts(base, sub, sup):
    if sub is not None and sup is not None:
        tag = 'sSubSup'
    elif sup is not None:
        tag = 'sSup'
    else:
        tag = 'sSub'
    element = _m(tag)
    element.append(_group(_convert(base) if base is not None else []))
    if sub is not None:
        element.append(_slot('sub', _convert(sub)))
    if sup is not None:
        element.append(_slot('sup', _convert(sup)))
    return element


def _nary_from_script(node, operand):
    kids = _children(node)
    tag = _local(node)
    base = kids[0] if kids else None
    sub = sup = None
    if tag == 'msub' and len(kids) > 1:
        sub = kids[1]
    elif tag == 'msup' and len(kids) > 1:
        sup = kids[1]
    elif tag == 'msubsup':
        sub = kids[1] if len(kids) > 1 else None
        sup = kids[2] if len(kids) > 2 else None
    return _nary(base, sub, sup, operand)


def _nary(base, sub, sup, operand):
    element = _m('nary')
    props = _m('naryPr')
    char = _m('chr')
    char.set(qn('m:val'), _operator_char(base) or '∑')
    props.append(char)
    location = _m('limLoc')
    location.set(qn('m:val'), 'undOvr')
    props.append(location)
    if sub is None:
        hide = _m('subHide')
        hide.set(qn('m:val'), '1')
        props.append(hide)
    if sup is None:
        hide = _m('supHide')
        hide.set(qn('m:val'), '1')
        props.append(hide)
    element.append(props)
    element.append(_slot('sub', _convert(sub) if sub is not None else []))
    element.append(_slot('sup', _convert(sup) if sup is not None else []))
    element.append(_group(operand))
    return element


def _accent(node):
    kids = _children(node)
    base = kids[0] if kids else None
    accent = kids[1] if len(kids) > 1 else None
    char = _text(accent) if accent is not None else '¯'
    if not char or char in _OVERLINE:
        char = '¯'
    element = _m('acc')
    props = _m('accPr')
    marker = _m('chr')
    marker.set(qn('m:val'), char[:1])
    props.append(marker)
    element.append(props)
    element.append(_group(_convert(base) if base is not None else []))
    return element


def _matrix(node):
    rows = [child for child in _children(node) if _local(child) == 'mtr']
    columns = 1
    if rows:
        columns = max(1, sum(1 for cell in _children(rows[0]) if _local(cell) == 'mtd'))
    element = _m('m')
    props = _m('mPr')
    columns_el = _m('mcs')
    column = _m('mc')
    column_props = _m('mcPr')
    count = _m('count')
    count.set(qn('m:val'), str(columns))
    align = _m('mcJc')
    align.set(qn('m:val'), 'left')
    column_props.append(count)
    column_props.append(align)
    column.append(column_props)
    columns_el.append(column)
    props.append(columns_el)
    element.append(props)
    for row in rows:
        row_el = _m('mr')
        for cell in _children(row):
            if _local(cell) != 'mtd':
                continue
            row_el.append(_group(_convert_row(_children(cell))))
        element.append(row_el)
    return element


def _delimiter(begin: str, end: str, elements):
    element = _m('d')
    props = _m('dPr')
    begin_el = _m('begChr')
    begin_el.set(qn('m:val'), begin if begin is not None else '')
    end_el = _m('endChr')
    end_el.set(qn('m:val'), end if end is not None else '')
    props.append(begin_el)
    props.append(end_el)
    element.append(props)
    element.append(_group(elements))
    return element


def _operator_char(node) -> str:
    if node is None:
        return ''
    tag = _local(node)
    if tag == 'mrow':
        kids = _children(node)
        return _operator_char(kids[0]) if kids else ''
    if tag == 'mo':
        return _text(node)
    return ''


def _is_relation(node) -> bool:
    if _local(node) != 'mo':
        return False
    return _text(node) in _RELATIONS
