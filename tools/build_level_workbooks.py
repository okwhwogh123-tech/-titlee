# -*- coding: utf-8 -*-
"""
기본부문(380) 투입산출표를 한국은행 '2020 기준년 상품분류(투입산출표)'에 따라
소분류(165)·중분류(83)로 합산해, 한은 배포 파일과 같은 12개 시트 형식의 엑셀을 만든다.

사용법 (저장소 최상위에서):
    python tools/build_level_workbooks.py

- 거래표(총·국산·수입)는 부문 행·열을 그대로 합산(가법성)한다.
- 계수표·유발계수표·최종수요 항목별 유발액은 합산된 거래표로 다시 계산한다
  (유발계수는 합산 후 역행렬을 새로 구해야 하므로 단순 합산하면 틀린다).
- 같은 방법으로 만든 대분류가 한은 공식 대분류 파일과 일치하는지 마지막에 검증한다.
"""
import os
import sys

import numpy as np
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import model as M  # noqa: E402

DATA = os.path.join(ROOT, 'data')
TAG = '기본부문 합산'


def rd(path):
    with open(path, 'rb') as f:
        return f.read()


def additive(sheet, keyfun, order):
    """부문 코드 행/열을 상위 분류로 합산. 부문이 아닌 행/열(합계·최종수요 등)은 그대로 둔다."""
    rows, cols = {}, {}
    for _, c, n in sheet.rowlist:
        k = keyfun(c)
        rows.setdefault(k if k else c, n)
    for _, c, n in sheet.cols:
        k = keyfun(c)
        cols.setdefault(k if k else c, n)
    rk = [k for k in order if k in rows] + [k for k in rows if k not in order]
    ck = [k for k in order if k in cols] + [k for k in cols if k not in order]
    val = {}
    for i, rc, _ in sheet.rowlist:
        r = keyfun(rc) or rc
        row = sheet.rows[i]
        for j, cc, _ in sheet.cols:
            c = keyfun(cc) or cc
            v = row[j] if j < len(row) else None
            if v is None:
                continue
            val[(r, c)] = val.get((r, c), 0.0) + M.num(v)
    return rk, ck, val


def write_sheet(ws, title, unit, rk, ck, val, rname, cname, corner='상품\n    상품'):
    ws.append([title + f' [{TAG}]'])
    ws.append(['2023년'])
    ws.append(['생산자가격'])
    ws.append([f'단위 : {unit}'] if unit else [None])
    ws.append(['                                   ' + corner, None] + ck)
    ws.append([None, None] + [cname.get(c, c) for c in ck])
    def cell(v):
        if v is None:
            return None
        v = float(v)
        return v if np.isfinite(v) else None
    for r in rk:
        ws.append([r, rname.get(r, r)] + [cell(val.get((r, c))) for c in ck])


def build(level, basic_sheets, roles, basic_model):
    m = M.aggregate_from('basic', level)
    codes, names = m.codes, m.names
    key = {b: M.parent_codes('basic', b).get(level) for b in basic_model.codes}
    keyfun = key.get
    nm = dict(zip(codes, names))
    items = [c for c, _ in M.ITEMS]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1) 거래표 3종: 가법 합산
    tnames = {}
    for role in ('Z', 'Zm', 'Zd'):
        src = basic_sheets[roles[role]]
        sh = M.Sheet(roles[role], src)
        rk, ck, val = additive(sh, keyfun, codes)
        rname = {**{c: n for _, c, n in sh.rowlist}, **nm}
        cname = {**{c: n for _, c, n in sh.cols}, **nm}
        tnames[role] = (rname, cname)
        write_sheet(wb.create_sheet(roles[role]), sh.meta[0] if sh.meta else roles[role],
                    '백만원', rk, ck, val, rname, cname)
    rname_z, cname_z = tnames['Z']

    # 2) 투입계수 3종: 재계산
    x = m.x
    X = x.sum()
    zs = M.Sheet(roles['Z'], basic_sheets[roles['Z']])
    _, _, zval = additive(zs, keyfun, codes)
    tot_rows = ['9590'] + [c for c, _ in M.VA_ITEMS] + ['9690', '9790']
    for role, mat, extra_rows, label in (('A', m.Ad + m.Am, tot_rows, '총투입계수표'),
                                         ('Am', m.Am, ['9590'], '수입투입계수표'),
                                         ('Ad', m.Ad, ['9590'], '국산투입계수표')):
        val = {}
        for a, r in enumerate(codes):
            for b, c in enumerate(codes):
                val[(r, c)] = mat[a, b]
            src = {'A': m.Zd + m.Zm, 'Am': m.Zm, 'Ad': m.Zd}[role]
            val[(r, '9090')] = src[a].sum() / X
        colsum = mat.sum(0)
        val[('9590', '9090')] = {'A': (m.Zd.sum() + m.Zm.sum()), 'Am': m.Zm.sum(), 'Ad': m.Zd.sum()}[role] / X
        for b, c in enumerate(codes):
            val[('9590', c)] = colsum[b]
        if role == 'A':
            for k, (vc, _) in enumerate(M.VA_ITEMS):
                for b, c in enumerate(codes):
                    val[(vc, c)] = M.safe_div(m.VA[k, b], x[b])
                val[(vc, '9090')] = m.VA[k].sum() / X
            for b, c in enumerate(codes):
                val[('9690', c)] = m.v[b]
                val[('9790', c)] = 1.0 if x[b] else 0.0
            val[('9690', '9090')] = m.va.sum() / X
            val[('9790', '9090')] = 1.0
        rlab = {'9590': '중간투입계' if role == 'A' else '소계'}
        ws = wb.create_sheet(roles[role])
        write_sheet(ws, f'투입산출표_{label}(2023년 가격)', None, codes + extra_rows,
                    codes + ['9090'], val, {**rname_z, **nm, **rlab}, {**nm, '9090': '중간수요계'})

    # 3) 유발계수 3종: 재계산 (+ 행합·열합, 영향력·감응도계수)
    for role, mat, label in (('L', m.L, '생산유발계수표'), ('M', m.M, '수입유발계수표'),
                             ('V', m.V, '부가가치유발계수표')):
        val = {}
        rs, cs = mat.sum(1), mat.sum(0)
        for a, r in enumerate(codes):
            for b, c in enumerate(codes):
                val[(r, c)] = mat[a, b]
            val[(r, '9090')] = rs[a]
        for b, c in enumerate(codes):
            val[('9590', c)] = cs[b]
        val[('9590', '9090')] = mat.sum()
        rows_, cols_ = codes + ['9590'], codes + ['9090']
        rlab, clab = {'9590': '열합'}, {'9090': '행합'}
        if role == 'L':
            bl, fl = cs / cs.mean(), rs / rs.mean()
            for b, c in enumerate(codes):
                val[('9801', c)] = bl[b]
            for a, r in enumerate(codes):
                val[(r, '9802')] = fl[a]
            val[('9801', '9090')] = bl.sum()
            val[('9590', '9802')] = fl.sum()
            rows_, cols_ = rows_ + ['9801'], cols_ + ['9802']
            rlab['9801'], clab['9802'] = '영향력계수', '감응도계수'
        write_sheet(wb.create_sheet(roles[role]), f'투입산출표_{label}(2023년 가격)', None,
                    rows_, cols_, val, {**nm, **rlab}, {**nm, **clab})

    # 4) 최종수요 항목별 유발액 3종: 재계산
    inames = {c: n for _, c, n in M.Sheet(roles['FP'], basic_sheets[roles['FP']]).cols}
    for role, mat, denom, label in (
            ('FP', m.L @ m.Fd, m.Fd, '생산유발액'),
            ('FM', m.M @ m.Fd + m.Fm, m.Fd + m.Fm, '수입유발액'),
            ('FV', m.V @ m.Fd, m.Fd, '부가가치유발액')):
        val = {}
        full = np.hstack([mat, mat.sum(1, keepdims=True)])
        den = np.append(denom.sum(0), denom.sum())
        cols_ = items + ['9190']
        for a, r in enumerate(codes):
            for b, c in enumerate(cols_):
                val[(r, c)] = full[a, b]
        tot = full.sum(0)
        for b, c in enumerate(cols_):
            val[('9819', c)] = tot[b]
            val[('9812', c)] = tot[b] / tot[-1] * 100
            val[('9811', c)] = tot[b] / den[b]
        write_sheet(wb.create_sheet(roles[role]), f'투입산출표_최종수요 항목별 {label}(2023년 가격)',
                    '백만원', codes + ['9819', '9812', '9811'], cols_, val,
                    {**nm, '9819': '열합', '9812': '구성비(%)', '9811': '유발계수'}, inames,
                    corner='최종수요 항목\n    상품')
    return wb, m


def main():
    M.load_classification(rd(os.path.join(DATA, 'classification.xlsx')))
    raw = rd(os.path.join(DATA, 'io-2023-basic.xlsx'))
    sheets, roles = M.parse_workbook(raw)
    M.load_level('basic', raw)
    basic = M.DATA['basic']

    # 검증: 같은 방법으로 만든 대분류가 한은 공식 대분류와 같은가
    wb, _ = build('large', sheets, roles, basic)
    path = os.path.join('/tmp', 'io-2023-large-check.xlsx')
    wb.save(path)
    M.load_level('large', rd(os.path.join(DATA, 'io-2023-large.xlsx')))
    official = M.DATA['large']
    gen = M.build_io('large', *M.parse_workbook(rd(path)))
    print('대분류 검증: 총산출 최대차', float(np.abs(gen.x - official.x).max()),
          '/ 생산유발계수 최대차', float(np.abs(gen.L - official.L).max()))

    for level, fname in (('small', 'io-2023-small.xlsx'), ('medium', 'io-2023-medium.xlsx')):
        wb, m = build(level, sheets, roles, basic)
        out = os.path.join(DATA, fname)
        wb.save(out)
        print(f'{M.LEVELS[level]["label"]}: {m.n}개 부문 → {out}')


if __name__ == '__main__':
    main()
