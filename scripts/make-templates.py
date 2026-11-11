# 기본 내장 양식 만들기: python3 scripts/make-templates.py   (openpyxl 필요)
# samples/합의서_기본양식.xlsx · samples/인보이스_기본양식.xlsx 를 만들고,
# 브라우저가 file:// 로 열려도 쓸 수 있게 js/builtin-templates.js 에 base64 로 넣습니다.
# 양식은 가상입니다 — 실제 회사 양식을 받으면(기획서 10장 ②) 그 양식에 {{키}} 만 적어 올리면 됩니다.
import base64
import io
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parent.parent
FONT = '맑은 고딕'
thin = Side(style='thin', color='7F8C9D')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill('solid', fgColor='E6ECF4')
LABEL_FILL = PatternFill('solid', fgColor='F3F6FA')
MONEY = '#,##0'


def f(size=10, bold=False):
    return Font(name=FONT, size=size, bold=bold)


def put(ws, ref, value, size=10, bold=False, align='left', fill=None, box=False, fmt=None, wrap=False):
    c = ws[ref]
    c.value = value
    c.font = f(size, bold)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if fill:
        c.fill = fill
    if box:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


def box_range(ws, rng, fill=None):
    for row in ws[rng]:
        for c in row:
            c.border = BOX
            if fill:
                c.fill = fill


def merge(ws, rng, value, **kw):
    ws.merge_cells(rng)
    put(ws, rng.split(':')[0], value, **kw)


def page(ws, area, landscape=False):
    ws.print_area = area
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    ws.sheet_view.showGridLines = False


def agreement():
    wb = Workbook()
    ws = wb.active
    ws.title = '합의서'
    for col, w in zip('ABCDEFGHI', [6, 15, 11, 11, 18, 7, 13, 9, 14]):
        ws.column_dimensions[col].width = w
    merge(ws, 'A1:I1', '클레임 보상 합의서', size=20, bold=True, align='center')
    ws.row_dimensions[1].height = 36
    rows = [
        (3, '문서 번호', '{{문서번호}}', '작성일', '{{발행일_한글}}'),
        (4, '갑(발주사)', '{{발행처}}', '을(협력사)', '{{협력사명}}'),
        (5, '담당 부서', '{{발행부서}}', '협력사 코드', '{{협력사코드}}'),
    ]
    for r, l1, v1, l2, v2 in rows:
        merge(ws, f'A{r}:B{r}', l1, bold=True, fill=LABEL_FILL)
        merge(ws, f'C{r}:E{r}', v1)
        merge(ws, f'F{r}:G{r}', l2, bold=True, fill=LABEL_FILL)
        merge(ws, f'H{r}:I{r}', v2)
        box_range(ws, f'A{r}:I{r}')
    merge(ws, 'A7:I7', '갑과 을은 아래 클레임의 보상 금액에 대하여 다음과 같이 합의한다.', size=11)
    merge(ws, 'A9:B9', '합의 금액', size=12, bold=True, fill=LABEL_FILL)
    merge(ws, 'C9:I9', '{{총액_표기}}', size=13, bold=True)
    box_range(ws, 'A9:I9')
    ws.row_dimensions[9].height = 26
    merge(ws, 'A10:I10', '(합의합계 {{합의합계}} · 부가세 {{부가세}} · 클레임 {{클레임건수}}건 · 청구합계 대비 {{부담비율}})', size=9)
    put(ws, 'A12', '1. 대상 클레임', size=11, bold=True)
    heads = ['번호', '클레임 번호', '발생일', '품번', '품명', '수량', '청구금액', '부담 비율', '합의금액']
    for i, h in enumerate(heads):
        put(ws, f'{"ABCDEFGHI"[i]}13', h, bold=True, align='center', fill=HEAD_FILL, box=True)
    rep = ['{{행.번호}}', '{{행.클레임번호}}', '{{행.발생일}}', '{{행.품번}}', '{{행.품명}}', '{{행.수량}}', '{{행.청구금액}}', '{{행.부담비율}}', '{{행.합의금액}}']
    aligns = ['center', 'left', 'center', 'left', 'left', 'right', 'right', 'right', 'right']
    fmts = [None, None, None, None, None, '#,##0', MONEY, None, MONEY]
    # 14행 = 반복 줄, 15~21행 = 미리 그린 빈 표 줄(모두 8줄). 넘치면 도구가 줄을 끼워 넣습니다.
    for r in range(14, 22):
        for i in range(9):
            put(ws, f'{"ABCDEFGHI"[i]}{r}', rep[i] if r == 14 else None, align=aligns[i], box=True, fmt=fmts[i])
    merge(ws, 'A22:E22', '합계', bold=True, align='center', fill=HEAD_FILL)
    put(ws, 'F22', '{{수량합계}}', bold=True, align='right', fill=HEAD_FILL, fmt='#,##0')
    put(ws, 'G22', '{{청구합계}}', bold=True, align='right', fill=HEAD_FILL, fmt=MONEY)
    put(ws, 'H22', '{{부담비율}}', bold=True, align='right', fill=HEAD_FILL)
    put(ws, 'I22', '=SUM(I14:I21)', bold=True, align='right', fill=HEAD_FILL, fmt=MONEY)
    box_range(ws, 'A22:I22')
    put(ws, 'A24', '2. 합의 조건', size=11, bold=True)
    clauses = [
        '① 을은 위 클레임이 을이 납품한 부품의 품질 문제로 발생하였음을 확인하고, 합의 금액을 갑에게 보상한다.',
        '② 을은 갑이 발행한 인보이스에 따라 {{지급기한}}까지 합의 금액을 지급한다.',
        '③ 이 합의로 위 클레임에 관한 보상은 끝난 것으로 하며, 같은 원인으로 새로 생기는 클레임은 따로 협의한다.',
    ]
    for i, t in enumerate(clauses):
        merge(ws, f'A{25 + i}:I{25 + i}', t, wrap=True)
        ws.row_dimensions[25 + i].height = 30
    merge(ws, 'A29:I29', '이 합의를 증명하기 위하여 합의서 2부를 작성하여 갑과 을이 서명(날인)한 뒤 1부씩 보관한다.', size=9)
    merge(ws, 'A31:I31', '{{합의일_한글}}', size=11, align='center')
    merge(ws, 'A33:D33', '갑: {{발행처}}', size=11, bold=True)
    merge(ws, 'F33:I33', '을: {{협력사명}}', size=11, bold=True)
    merge(ws, 'A34:D34', '{{발행부서}} {{담당자}}            (인)')
    merge(ws, 'F34:I34', '대표자                         (인)')
    page(ws, 'A1:I34')
    return wb


def invoice():
    wb = Workbook()
    ws = wb.active
    ws.title = '인보이스'
    for col, w in zip('ABCDEFGH', [6, 15, 11, 11, 22, 7, 7, 15]):
        ws.column_dimensions[col].width = w
    merge(ws, 'A1:H1', 'INVOICE (클레임 보상 청구서)', size=18, bold=True, align='center')
    ws.row_dimensions[1].height = 34
    info = [
        (3, 'Invoice No.', '{{문서번호}}', 'Date', '{{발행일}}'),
        (4, '관련 합의서', '{{관련문서번호}}', '지급 기한', '{{지급기한}}'),
    ]
    for r, l1, v1, l2, v2 in info:
        merge(ws, f'A{r}:B{r}', l1, bold=True, fill=LABEL_FILL)
        merge(ws, f'C{r}:E{r}', v1)
        merge(ws, f'F{r}:G{r}', l2, bold=True, fill=LABEL_FILL)
        put(ws, f'H{r}', v2)
        box_range(ws, f'A{r}:H{r}')
    parties = [
        (6, '청구처 (Bill To)', '{{협력사명}} ({{협력사코드}})'),
        (7, '발행처 (From)', '{{발행처}} {{발행부서}}'),
        (8, '담당 · 연락처', '{{담당자}} {{연락처}}'),
    ]
    for r, l, v in parties:
        merge(ws, f'A{r}:B{r}', l, bold=True, fill=LABEL_FILL)
        merge(ws, f'C{r}:H{r}', v)
        box_range(ws, f'A{r}:H{r}')
    heads = ['No', '클레임 번호', '발생일', '품번', '품명', '수량', '통화', '금액']
    for i, h in enumerate(heads):
        put(ws, f'{"ABCDEFGH"[i]}10', h, bold=True, align='center', fill=HEAD_FILL, box=True)
    rep = ['{{행.번호}}', '{{행.클레임번호}}', '{{행.발생일}}', '{{행.품번}}', '{{행.품명}}', '{{행.수량}}', '{{통화}}', '{{행.합의금액}}']
    aligns = ['center', 'left', 'center', 'left', 'left', 'right', 'center', 'right']
    # 11행 = 반복 줄, 12~17행 = 빈 표 줄(모두 7줄)
    for r in range(11, 18):
        for i in range(8):
            put(ws, f'{"ABCDEFGH"[i]}{r}', rep[i] if r == 11 else None, align=aligns[i], box=True, fmt=MONEY if i == 7 else None)
    sums = [(18, '공급가액 (합의합계)', '{{합의합계}}', False), (19, '부가세', '{{부가세}}', False), (20, '합계 (Total)', '{{총액}}', True)]
    for r, l, v, b in sums:
        merge(ws, f'A{r}:G{r}', l, bold=b, align='right', fill=LABEL_FILL)
        put(ws, f'H{r}', v, bold=b, align='right', fmt=MONEY)
        box_range(ws, f'A{r}:H{r}')
    merge(ws, 'A21:B21', '금액 (한글)', bold=True, fill=LABEL_FILL)
    merge(ws, 'C21:H21', '{{총액_표기}}', bold=True)
    box_range(ws, 'A21:H21')
    merge(ws, 'A23:B23', '입금 계좌', bold=True, fill=LABEL_FILL)
    merge(ws, 'C23:H23', '{{입금계좌}}')
    box_range(ws, 'A23:H23')
    merge(ws, 'A24:H24', '위 금액을 지급 기한까지 입금 계좌로 보내 주십시오.', size=9)
    merge(ws, 'E26:H26', '{{발행처}}', size=11, bold=True, align='right')
    merge(ws, 'E27:H27', '발행 담당: {{담당자}}        (인)', align='right')
    page(ws, 'A1:H27')
    return wb


def main():
    out = {}
    for key, name, make in [('ag', '합의서_기본양식.xlsx', agreement), ('inv', '인보이스_기본양식.xlsx', invoice)]:
        buf = io.BytesIO()
        make().save(buf)
        data = buf.getvalue()
        (ROOT / 'samples' / name).write_bytes(data)
        out[key] = (name, base64.b64encode(data).decode('ascii'))
        print('wrote samples/' + name, len(data), 'bytes')
    js = ['/* 기본 내장 양식(가상) — scripts/make-templates.py 가 만든 파일입니다. 직접 고치지 말고 스크립트를 고쳐 다시 만드세요. */',
          '(function (root) {', "  'use strict';",
          '  var T = {']
    for key, (name, b64) in out.items():
        js.append(f"    {key}: {{ name: '{name}', b64: '{b64}' }},")
    js[-1] = js[-1].rstrip(',')
    js += ['  };', "  if (typeof module === 'object' && module.exports) module.exports = T; else root.CLBuiltin = T;",
           "})(typeof self !== 'undefined' ? self : this);", '']
    (ROOT / 'js' / 'builtin-templates.js').write_text('\n'.join(js), encoding='utf-8')
    print('wrote js/builtin-templates.js')


if __name__ == '__main__':
    main()
