# -*- coding: utf-8 -*-
"""
2023년 산업연관표(연장표, 생산자가격) 분석 엔진
- 브라우저(Pyodide)와 일반 Python 양쪽에서 동일하게 동작한다.
- 모든 금액 단위: 백만원(당해년 가격). 계수: 무차원(원/원).
- 계수는 거래표에서 직접 다시 계산하고, 원표의 계수 시트와 대조해 검증한다.
"""
import io
import json
import re

import numpy as np
import openpyxl

# ---------------------------------------------------------------------------
# 0. 고정 정보: 최종수요 항목, 부가가치 항목, 분류 체계, 한은 공표치
# ---------------------------------------------------------------------------
ITEMS = [('9111', '민간소비'), ('9112', '정부소비'), ('9121', '민간고정자본형성'),
         ('9122', '정부고정자본형성'), ('9131', '재고증감'), ('9132', '귀중품순취득'),
         ('9140', '수출')]
ITEM_GROUP = {'9111': '소비', '9112': '소비', '9121': '투자', '9122': '투자',
              '9131': '투자', '9132': '투자', '9140': '수출'}
VA_ITEMS = [('9610', '피용자보수'), ('9621', '영업잉여'), ('9622', '고정자본소모'),
            ('9630', '생산세(보조금공제)')]

LEVELS = {
    'basic':  {'label': '기본부문', 'width': 4, 'official': True},
    'small':  {'label': '소분류', 'width': 3, 'official': True},
    'medium': {'label': '중분류', 'width': 2, 'official': True},
    'large':  {'label': '대분류', 'width': 2, 'official': True},
    's13':    {'label': '13부문', 'width': 2, 'official': False},
    's7':     {'label': '7부문', 'width': 1, 'official': False},
    's1':     {'label': '1부문(전산업)', 'width': 1, 'official': False},
}
LEVEL_ORDER = ['basic', 'small', 'medium', 'large', 's13', 's7', 's1']

# 대분류(33) ← 중분류(83) 코드 범위 (2020 기준년 상품분류, 2023 연장표 동일)
# 공식 부문분류표(data/classification.xlsx)를 불러오면 그 내용이 우선한다.
LARGE = [
    ('A', '농림수산품', 1, 5), ('B', '광산품', 6, 7), ('C01', '음식료품', 8, 10),
    ('C02', '섬유 및 가죽제품', 11, 12), ('C03', '목재 및 종이, 인쇄', 13, 15),
    ('C04', '석탄 및 석유제품', 16, 16), ('C05', '화학제품', 17, 24),
    ('C06', '비금속광물제품', 25, 26), ('C07', '1차 금속제품', 27, 29),
    ('C08', '금속가공제품', 30, 30), ('C09', '컴퓨터, 전자 및 광학기기', 31, 36),
    ('C10', '전기장비', 37, 37), ('C11', '기계 및 장비', 38, 39), ('C12', '운송장비', 40, 42),
    ('C13', '기타 제조업제품', 43, 43), ('C14', '제조임가공 및 산업용 장비 수리', 44, 44),
    ('D', '전력, 가스 및 증기', 45, 46), ('E', '수도, 폐기물처리 및 재활용서비스', 47, 49),
    ('F', '건설', 50, 51), ('G', '도소매 및 상품중개서비스', 52, 52),
    ('H', '운송서비스', 53, 57), ('I', '음식점 및 숙박서비스', 58, 58),
    ('J', '정보통신 및 방송서비스', 59, 64), ('K', '금융 및 보험서비스', 65, 67),
    ('L', '부동산서비스', 68, 69), ('M', '전문, 과학 및 기술 서비스', 70, 72),
    ('N', '사업지원서비스', 73, 74), ('O', '공공행정, 국방 및 사회보장', 75, 75),
    ('P', '교육서비스', 76, 76), ('Q', '보건 및 사회복지서비스', 77, 78),
    ('R', '예술, 스포츠 및 여가 관련 서비스', 79, 80), ('S', '기타 서비스', 81, 82),
    ('T', '기타', 83, 83)]
MEDIUM_TO_LARGE = {f'{m:02d}': code for code, _, a, b in LARGE for m in range(a, b + 1)}
LARGE_NAME = {code: name for code, name, _, _ in LARGE}

# 13부문·7부문: 한은 보도자료의 부문 구성(대분류 기준 재구성)
S13 = [('01', '농림수산품', ['A']), ('02', '광산품', ['B']),
       ('03', '소비재제품', ['C01', 'C02', 'C13']),
       ('04', '기초소재제품', ['C03', 'C04', 'C05', 'C06', 'C07']),
       ('05', '조립가공제품', ['C08', 'C09', 'C10', 'C11', 'C12']),
       ('06', '제조임가공·산업용 장비 수리', ['C14']),
       ('07', '전력·가스·수도 및 폐기물', ['D', 'E']), ('08', '건설', ['F']),
       ('09', '도소매 및 운송', ['G', 'H']),
       ('10', '생산자서비스', ['J', 'K', 'L', 'M', 'N']),
       ('11', '사회서비스', ['O', 'P', 'Q']),
       ('12', '소비자서비스', ['I', 'R', 'S']), ('13', '기타', ['T'])]
LARGE_TO_S13 = {l: c for c, _, ls in S13 for l in ls}
S7 = [('1', '농림수산품', ['01']), ('2', '광산품', ['02']),
      ('3', '공산품', ['03', '04', '05', '06']), ('4', '전력·가스·수도 및 폐기물', ['07']),
      ('5', '건설', ['08']), ('6', '서비스', ['09', '10', '11', '12']), ('7', '기타', ['13'])]
S13_TO_S7 = {s: c for c, _, ss in S7 for s in ss}

# 공식 부문분류표(기본→소→중→대) — load_classification()이 채운다
CLASS = {'basic': {}, 'names': {'basic': {}, 'small': {}, 'medium': {}, 'large': {}}}

# 한국은행 보도자료(2025.9.24, 2023년 산업연관표 연장표) 공표치
BOK = {
    'economy': [
        ('총산출', 'x_total', 5786.7, '조원'),
        ('총공급(=총수요)', 'supply_total', 6802.7, '조원'),
        ('수입', 'import_total', 1016.0, '조원'),
        ('부가가치율', 'va_rate', 41.2, '%'),
        ('중간투입률', 'mid_rate', 58.8, '%'),
        ('국산 중간투입률', 'dom_mid_rate', 45.2, '%'),
        ('수입의존도(수입 중간투입률)', 'imp_dep', 13.6, '%'),
        ('국산화율', 'localization', 77.0, '%'),
        ('수출률(수출/총산출)', 'export_rate', 17.3, '%'),
        ('수입률(수입/총공급)', 'import_rate', 14.9, '%'),
        ('생산유발계수(전체)', 'prod_mult', 1.827, '배'),
        ('부가가치유발계수(전체)', 'va_mult', 0.752, '배'),
        ('수입유발계수(전체, 부문별 기준)', 'imp_mult', 0.248, '배'),
        ('수입유발계수(최종수요 항목 기준)', 'imp_mult_fd', 0.299, '배'),
        ('생산유발 중 소비 비중', 'share_cons_prod', 44.2, '%'),
        ('생산유발 중 수출 비중', 'share_exp_prod', 32.7, '%'),
        ('부가가치유발 중 소비 비중', 'share_cons_va', 52.1, '%'),
        ('부가가치유발 중 수출 비중', 'share_exp_va', 25.4, '%'),
    ],
    # 최종수요 항목별 유발계수(생산, 부가가치, 수입) — 한은은 대분류(33) 표로 계산해 공표
    'fd': {'9111': (1.787, 0.810, 0.283), '9112': (1.529, 0.885, 0.115),
           '9121': (1.956, 0.793, 0.306), '9122': (2.036, 0.787, 0.277),
           '9131': (2.053, 0.649, 0.375), '9140': (1.901, 0.608, 0.394),
           '소비': (1.713, 0.831, 0.239), '투자': (1.969, 0.790, 0.303), '수출': (1.901, 0.608, 0.394)},
    's7': {'1': (1.936, 0.805, 0.195), '2': (2.055, 0.832, 0.168), '3': (1.975, 0.591, 0.409),
           '4': (1.638, 0.487, 0.513), '5': (2.029, 0.788, 0.212), '6': (1.695, 0.857, 0.143)},
    's13': {'03': (2.125, 0.673, 0.327), '04': (1.807, 0.520, 0.480), '05': (2.056, 0.606, 0.394),
            '06': (1.955, 0.852, 0.148), '09': (1.737, 0.818, 0.182), '10': (1.656, 0.888, 0.113),
            '11': (1.541, 0.880, 0.121), '12': (2.068, 0.802, 0.198)},
}

# 시트 역할 판별(시트명 키워드)과 설명
ROLES = [
    ('Z', '총거래', '총거래표', '국산품과 수입품을 합친 거래액. 행은 상품의 판매처(중간수요·최종수요), 열은 산업의 투입 구조(중간투입·부가가치).', '백만원'),
    ('Zm', '수입거래', '수입거래표', '수입품이 어느 산업의 중간재·어떤 최종수요로 쓰였는지 기록.', '백만원'),
    ('Zd', '국산거래', '국산거래표', '국내에서 생산된 상품의 거래만 기록. 레온티에프 역행렬(국산)의 기초.', '백만원'),
    ('A', '총투입계수', '총투입계수', '열 산업이 1원어치를 생산할 때 투입하는 각 상품(국산+수입)·부가가치의 크기. a_ij = z_ij / x_j', '계수(무차원, 원/원)'),
    ('Am', '수입투입계수', '수입투입계수', '산출 1원당 수입 중간재 투입액.', '계수(무차원, 원/원)'),
    ('Ad', '국산투입계수', '국산투입계수', '산출 1원당 국산 중간재 투입액. Ad', '계수(무차원, 원/원)'),
    ('L', '생산유발계수', '생산유발계수', '(I-Ad)^-1. 열 상품의 최종수요 1원이 각 행 산업에서 직·간접으로 유발하는 생산액. 열합=부문별 생산유발계수, 영향력·감응도계수 포함.', '계수(무차원, 원/원)'),
    ('M', '수입유발계수', '수입유발계수', 'Am·(I-Ad)^-1. 최종수요 1원이 유발하는 상품별 수입액.', '계수(무차원, 원/원)'),
    ('V', '부가가치유발계수', '부가가치유발계수', 'diag(v)·(I-Ad)^-1. 최종수요 1원이 유발하는 산업별 부가가치.', '계수(무차원, 원/원)'),
    ('FP', '생산유발액', '최종수요 항목별 생산유발액', '실제 최종수요(소비·투자·수출)가 유발한 산업별 생산액. 하단에 열합·구성비(%)·유발계수.', '백만원'),
    ('FM', '수입유발액', '최종수요 항목별 수입유발액', '최종수요 항목별로 유발된 상품별 수입액(최종재 직접 수입 포함).', '백만원'),
    ('FV', '부가가치유발액', '최종수요 항목별 부가가치유발액', '최종수요 항목별로 유발된 산업별 부가가치.', '백만원'),
]
ROW_UNITS = {'구성비': '%', '유발계수': '계수(무차원)', '영향력계수': '지수(전산업 평균=1)',
             '감응도계수': '지수(전산업 평균=1)'}

DATA = {}    # level -> IO 객체
BOOKS = {}   # level -> {'sheets': {...}, 'roles': {...}}


# ---------------------------------------------------------------------------
# 1. 유틸리티
# ---------------------------------------------------------------------------
def emit(obj):
    return json.dumps(_clean(obj), ensure_ascii=False, allow_nan=False)


def _clean(o):
    if isinstance(o, dict):
        return {k: _clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_clean(v) for v in o]
    if isinstance(o, (np.floating, float)):
        f = float(o)
        return f if np.isfinite(f) else None
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.ndarray):
        return _clean(o.tolist())
    return o


def norm(c):
    """셀 값 → 코드 문자열. 숫자로 저장된 코드도 문자열로 통일."""
    if c is None:
        return None
    if isinstance(c, bool):
        return None
    if isinstance(c, (int, np.integer)):
        return str(int(c))
    if isinstance(c, float):
        return str(int(c)) if c.is_integer() else None
    s = str(c).strip()
    return s or None


CODE_RE = re.compile(r'\d{1,4}|[A-Z]\d{0,2}')


def is_code(c):
    s = norm(c)
    return bool(s) and CODE_RE.fullmatch(s) is not None


def clean_name(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def is_total_code(code):
    return len(code) == 4 and code[0] == '9'


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    try:
        return float(str(v).replace(',', ''))
    except ValueError:
        return float('nan')


def safe_div(a, b):
    a = np.asarray(a, dtype=float)
    b = np.asarray(b, dtype=float)
    out = np.zeros(np.broadcast(a, b).shape)
    np.divide(a, b, out=out, where=(b != 0))
    return out


# ---------------------------------------------------------------------------
# 2. 엑셀 파싱
# ---------------------------------------------------------------------------
class Sheet:
    """한은 투입산출표 시트 1개를 표준 구조(헤더행·명칭행·데이터행)로 해석."""

    def __init__(self, name, rows):
        self.name = name
        self.rows = rows
        h = None
        for i, r in enumerate(rows[:20]):
            if r and sum(1 for c in r[2:] if is_code(c)) >= 2:
                h = i
                break
        if h is None:
            raise ValueError(f'[{name}] 상품 코드가 있는 헤더 행을 찾지 못했습니다.')
        self.h = h
        nrow = rows[h + 1] if h + 1 < len(rows) else ()
        has_name_row = bool(nrow) and norm(nrow[0]) is None
        self.data_start = h + 2 if has_name_row else h + 1
        self.meta = [clean_name(r[0]) for r in rows[:h] if r and r[0] is not None]
        self.cols = []
        for j, c in enumerate(rows[h]):
            if j < 2 or not is_code(c):
                continue
            nm = nrow[j] if has_name_row and j < len(nrow) else ''
            self.cols.append((j, norm(c), clean_name(nm)))
        self.rowlist = []
        for i in range(self.data_start, len(rows)):
            r = rows[i]
            if r and is_code(r[0]):
                self.rowlist.append((i, norm(r[0]), clean_name(r[1] if len(r) > 1 else '')))
        self.ci = {c: j for j, c, _ in self.cols}
        self.ri = {c: i for i, c, _ in self.rowlist}
        self.unit = next((m.split(':', 1)[-1].strip() for m in self.meta if '단위' in m), None)

    def get(self, r, c):
        row = self.rows[self.ri[r]]
        j = self.ci[c]
        return num(row[j]) if j < len(row) else 0.0

    def matrix(self, rcodes, ccodes):
        out = np.zeros((len(rcodes), len(ccodes)))
        for a, r in enumerate(rcodes):
            row = self.rows[self.ri[r]]
            for b, c in enumerate(ccodes):
                j = self.ci[c]
                out[a, b] = num(row[j]) if j < len(row) else 0.0
        return out

    def has_row(self, c):
        return c in self.ri

    def has_col(self, c):
        return c in self.ci


def find_code(sheet, candidates, axis, name_keys=()):
    """코드 우선, 없으면 명칭 키워드로 총계 행/열을 찾는다."""
    pool = sheet.rowlist if axis == 'row' else sheet.cols
    codes = {c for _, c, _ in pool}
    for c in candidates:
        if c in codes:
            return c
    for _, c, nm in pool:
        if any(k in nm.replace(' ', '') for k in name_keys):
            return c
    return None


def detect_roles(sheetnames):
    roles = {}
    for s in sheetnames:
        key = s.replace(' ', '')
        for role, kw, *_ in ROLES:
            if kw in key and role not in roles:
                roles[role] = s
                break
    return roles


# ---------------------------------------------------------------------------
# 3. 산업연관 모형 (IO)
# ---------------------------------------------------------------------------
class IO:
    """거래표(국산·수입)로부터 모든 계수와 유발효과를 계산하는 모형."""

    def __init__(self, level, codes, names, Zd, Zm, x, Fd, Fm, VA, imports, source):
        self.level, self.codes, self.names = level, list(codes), list(names)
        self.n = len(codes)
        self.Zd, self.Zm, self.x = Zd, Zm, x
        self.Fd, self.Fm, self.VA = Fd, Fm, VA        # Fd,Fm: n×7, VA: 4×n
        self.imports = imports                        # n (수입계: 수입+생산물세)
        self.source = source                          # 'official' | 'aggregated:<from>'
        self.active = x > 0
        xs = np.where(self.active, x, 1.0)
        self.Ad = Zd / xs
        self.Am = Zm / xs
        self.Ad[:, ~self.active] = 0.0
        self.Am[:, ~self.active] = 0.0
        self.va = self.VA.sum(0)
        self.v = safe_div(self.va, x)
        self.w = safe_div(self.VA[0], x)
        self.L = np.linalg.inv(np.eye(self.n) - self.Ad)
        self.M = self.Am @ self.L
        self.V = self.v[:, None] * self.L
        self.idx = {c: i for i, c in enumerate(self.codes)}
        self.fd = Fd.sum(1)

    # 분류 집계 ---------------------------------------------------------
    def aggregate(self, level, keys, key_names, source):
        groups = sorted(set(keys))
        gi = {g: i for i, g in enumerate(groups)}
        S = np.zeros((len(groups), self.n))
        for j, k in enumerate(keys):
            S[gi[k], j] = 1.0
        return IO(level, groups, [key_names.get(g, g) for g in groups],
                  S @ self.Zd @ S.T, S @ self.Zm @ S.T, S @ self.x,
                  S @ self.Fd, S @ self.Fm, self.VA @ S.T, S @ self.imports, source)


def parent_codes(level, code):
    """해당 코드의 상위 분류 코드들(공식 부문분류표 우선, 없으면 코드 앞자리 규칙)."""
    out = {}
    if level in ('basic', 'small', 'medium'):
        c = code.zfill(LEVELS[level]['width'])
        if level == 'basic' and c in CLASS['basic']:
            out.update(CLASS['basic'][c])
        else:
            if level == 'basic':
                out['small'] = c[:3]
            out['medium'] = c[:2]
            out['large'] = MEDIUM_TO_LARGE.get(c[:2])
    elif level == 'large':
        out['large'] = code
    elif level == 's13':
        return {'s13': code, 's7': S13_TO_S7.get(code), 's1': '1'}
    elif level == 's7':
        return {'s7': code, 's1': '1'}
    else:
        return {'s1': '1'}
    out['s13'] = LARGE_TO_S13.get(out.get('large'))
    out['s7'] = S13_TO_S7.get(out['s13']) if out['s13'] else None
    out['s1'] = '1'
    return out


def load_classification(raw):
    """한은 '2020 기준년 상품분류(투입산출표)' 엑셀: 기본/소/중/대 코드·명칭 열(병합셀은 앞값 채움)."""
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    rows = list(wb[wb.sheetnames[0]].values)
    wb.close()
    keys = ['basic', 'small', 'medium', 'large']
    cur = [None] * 4
    basic = {}
    names = {k: {} for k in keys}
    for r in rows:
        if not r or not is_code(r[0]) or is_total_code(norm(r[0])):
            continue
        for k in range(4):
            c = norm(r[2 * k]) if 2 * k < len(r) else None
            if c and is_code(c):
                cur[k] = c
                names[keys[k]][c] = clean_name(r[2 * k + 1])
        basic[cur[0]] = {'small': cur[1], 'medium': cur[2], 'large': cur[3]}
    if len(basic) < 100:
        raise ValueError('부문분류표 형식을 인식하지 못했습니다.')
    CLASS['basic'] = basic
    CLASS['names'] = names
    for b, p in basic.items():
        MEDIUM_TO_LARGE[p['medium']] = p['large']
    LARGE_NAME.update(names['large'])
    return emit({'counts': {k: len(v) for k, v in names.items()}})


def group_name(level, code):
    if level in DATA and code in DATA[level].idx:
        return DATA[level].names[DATA[level].idx[code]]
    if code in CLASS['names'].get(level, {}):
        return CLASS['names'][level][code]
    if level == 'large':
        return LARGE_NAME.get(code, code)
    if level == 's13':
        return next((n for c, n, _ in S13 if c == code), code)
    if level == 's7':
        return next((n for c, n, _ in S7 if c == code), code)
    if level == 's1':
        return '전산업'
    return code


def parse_workbook(raw):
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    sheets = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    roles = detect_roles(list(sheets))
    missing = [r for r in ('Z', 'Zd', 'Zm') if r not in roles]
    if missing:
        raise ValueError('필수 시트(총·국산·수입 거래표)를 찾지 못했습니다: ' + ', '.join(missing))
    return sheets, roles


def build_io(level, sheets, roles):
    Zs = Sheet(roles['Z'], sheets[roles['Z']])
    Zd = Sheet(roles['Zd'], sheets[roles['Zd']])
    Zm = Sheet(roles['Zm'], sheets[roles['Zm']])
    rowset = {c for _, c, _ in Zs.rowlist}
    codes = [c for _, c, _ in Zs.cols if c in rowset and not is_total_code(c)
             and Zd.has_row(c) and Zd.has_col(c) and Zm.has_row(c)]
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('산업 코드를 판별할 수 없습니다.')
    nm = {c: n for _, c, n in Zs.cols}
    names = [nm.get(c) or next(n for _, cc, n in Zs.rowlist if cc == c) for c in codes]
    xcol = find_code(Zs, ['9310'], 'col', ('총산출',))
    xrow = find_code(Zs, ['9790'], 'row', ('총투입',))
    if xcol:
        x = Zs.matrix(codes, [xcol])[:, 0]
    elif xrow:
        x = Zs.matrix([xrow], codes)[0]
    else:
        raise ValueError('총산출(9310) 또는 총투입계(9790)를 찾지 못했습니다.')
    items = [c for c, _ in ITEMS]
    for c in items:
        if not Zd.has_col(c):
            raise ValueError(f'최종수요 항목 {c} 열이 국산거래표에 없습니다.')
    VA = np.vstack([Zs.matrix([c], codes)[0] if Zs.has_row(c) else np.zeros(len(codes))
                    for c, _ in VA_ITEMS])
    impcol = find_code(Zs, ['9329'], 'col', ('수입계',))
    imports = Zs.matrix(codes, [impcol])[:, 0] if impcol else np.zeros(len(codes))
    generated = any('합산' in t for t in Zs.meta)
    model = IO(level, codes, names, Zd.matrix(codes, codes), Zm.matrix(codes, codes), x,
               Zd.matrix(codes, items), Zm.matrix(codes, items), VA, imports,
               'generated' if generated else 'official')
    return model


def verify(model, sheets, roles):
    """원표의 계수·유발액 시트를 거래표로 재계산한 값과 대조(최대 절대오차)."""
    res = {}
    codes = model.codes
    items = [c for c, _ in ITEMS]
    Zs = Sheet(roles['Z'], sheets[roles['Z']])
    tot = Zs.matrix(codes, codes) - model.Zd - model.Zm
    res['총거래표 = 국산 + 수입'] = float(np.max(np.abs(tot)))
    vtot = find_code(Zs, ['9690'], 'row', ('부가가치계',))
    if vtot:
        res['부가가치계 = 항목 합'] = float(np.max(np.abs(Zs.matrix([vtot], codes)[0] - model.va)))
    res['총산출 = L × 국산 최종수요'] = float(np.max(np.abs(model.L @ model.fd - model.x)) /
                                       max(model.x.max(), 1))
    checks = [('Ad', model.Ad, '국산투입계수'), ('Am', model.Am, '수입투입계수'),
              ('L', model.L, '생산유발계수'), ('M', model.M, '수입유발계수'),
              ('V', model.V, '부가가치유발계수')]
    for role, mat, label in checks:
        if role in roles:
            sh = Sheet(roles[role], sheets[roles[role]])
            if all(sh.has_row(c) and sh.has_col(c) for c in codes):
                res[label + ' (재계산 vs 원표)'] = float(np.max(np.abs(sh.matrix(codes, codes) - mat)))
    fd_checks = [('FP', model.L @ model.Fd, '생산유발액'),
                 ('FM', model.M @ model.Fd + model.Fm, '수입유발액'),
                 ('FV', model.V @ model.Fd, '부가가치유발액')]
    for role, mat, label in fd_checks:
        if role in roles:
            sh = Sheet(roles[role], sheets[roles[role]])
            if all(sh.has_row(c) for c in codes) and all(sh.has_col(c) for c in items):
                err = np.max(np.abs(sh.matrix(codes, items) - mat)) / max(np.abs(mat).max(), 1)
                res[label + ' (상대오차)'] = float(err)
    return res


# ---------------------------------------------------------------------------
# 4. JS에서 호출하는 함수들 (입력 JSON 문자열 → 출력 JSON 문자열)
# ---------------------------------------------------------------------------
def sector_list(m):
    out = []
    for i, c in enumerate(m.codes):
        out.append({'code': c, 'name': m.names[i], 'x': m.x[i], 'active': bool(m.active[i]),
                    'fd': [m.Fd[i, k] for k in range(len(ITEMS))],
                    'parents': {k: v for k, v in parent_codes(m.level, c).items() if v}})
    return out


def level_summary(level):
    m = DATA[level]
    return {'level': level, 'label': LEVELS[level]['label'], 'n': m.n, 'source': m.source,
            'sectors': sector_list(m), 'items': [{'code': c, 'name': n} for c, n in ITEMS]}


def level_summary_json(level):
    get_level(level)
    return emit(level_summary(level))


def load_level(level, raw):
    """공식 엑셀 1개를 읽어 해당 분류 수준 모형을 만든다."""
    sheets, roles = parse_workbook(raw)
    model = build_io(level, sheets, roles)
    expected = {'basic': 380, 'small': 165, 'medium': 83, 'large': 33}.get(level)
    warn = []
    if expected and model.n != expected:
        warn.append(f'{LEVELS[level]["label"]} 표준 부문 수({expected})와 다른 {model.n}개 부문이 감지되었습니다.')
    residuals = verify(model, sheets, roles)
    DATA[level] = model
    BOOKS[level] = {'sheets': sheets, 'roles': roles}
    derive_upper()
    out = level_summary(level)
    out.update({'residuals': residuals, 'warnings': warn, 'sheets': sheet_catalog(level),
                'derived': [l for l in ('s13', 's7', 's1') if l in DATA]})
    return emit(out)


def aggregate_from(src, level):
    """하위 분류에서 상위 분류를 합산해 만든다(공식 파일이 없을 때 대체, 검증용)."""
    m = DATA[src]
    keys = [parent_codes(src, c).get(level) for c in m.codes]
    if any(k is None for k in keys):
        raise ValueError(f'{LEVELS[src]["label"]} → {LEVELS[level]["label"]} 대응표가 없습니다.')
    names = {}
    for c, k in zip(m.codes, keys):
        names.setdefault(k, []).append(m.names[m.idx[c]])
    if level in ('small', 'medium', 'large') and CLASS['names'][level]:
        nm = {k: CLASS['names'][level].get(k, k) for k in names}
    elif level == 'large':
        nm = {k: LARGE_NAME.get(k, k) for k in names}
    elif level in ('s13', 's7', 's1'):
        nm = {k: group_name(level, k) for k in names}
    else:
        nm = {k: (v[0] + (f' 외 {len(v) - 1}개' if len(v) > 1 else '')) for k, v in names.items()}
    return m.aggregate(level, keys, nm, 'aggregated:' + src)


def derive_upper():
    """13·7·1부문은 대분류(공식 우선) 또는 기본부문에서 합산."""
    src = 'large' if 'large' in DATA and DATA['large'].source in ('official', 'generated') else (
        'basic' if 'basic' in DATA else None)
    if src is None:
        src = next((l for l in ('medium', 'small', 'large') if l in DATA), None)
    if src is None:
        return
    for lv in ('s13', 's7', 's1'):
        DATA[lv] = aggregate_from(src, lv)


def build_from_basic(level):
    if 'basic' not in DATA:
        raise ValueError('먼저 기본부문 자료를 불러오세요.')
    DATA[level] = aggregate_from('basic', level)
    BOOKS.pop(level, None)
    derive_upper()
    out = level_summary(level)
    out.update({'residuals': {}, 'warnings': ['공식 파일 대신 기본부문 거래표를 합산해 만든 자료입니다.'],
                'sheets': [], 'derived': ['s13', 's7', 's1']})
    return emit(out)


def get_level(level):
    if level not in DATA:
        raise ValueError(f'{LEVELS.get(level, {}).get("label", level)} 자료가 아직 준비되지 않았습니다.')
    return DATA[level]


# --- 원자료 탐색 ------------------------------------------------------------
def sheet_catalog(level):
    if level not in BOOKS:
        return []
    roles = BOOKS[level]['roles']
    inv = {v: k for k, v in roles.items()}
    out = []
    for s in BOOKS[level]['sheets']:
        role = inv.get(s)
        info = next((r for r in ROLES if r[0] == role), None)
        out.append({'name': s, 'role': role, 'title': info[2] if info else s,
                    'desc': info[3] if info else '', 'unit': info[4] if info else ''})
    return out


def explore(payload):
    p = json.loads(payload)
    level, sname, rowcode = p['level'], p['sheet'], p.get('row') or ''
    if level not in BOOKS:
        raise ValueError('이 분류 수준은 원본 엑셀이 없어 원자료 탐색을 할 수 없습니다.')
    rows = BOOKS[level]['sheets'][sname]
    sh = Sheet(sname, rows)
    choices = [{'code': c, 'name': n} for _, c, n in sh.rowlist]
    chosen = next((r for r in sh.rowlist if r[1] == rowcode), sh.rowlist[0])
    row = rows[chosen[0]]
    cells = []
    for j, c, n in sh.cols:
        v = row[j] if j < len(row) else None
        if isinstance(v, (int, float)):
            v = float(v)
        elif v is not None:
            v = str(v)
        cells.append({'code': c, 'name': n, 'value': v})
    cat = next((s for s in sheet_catalog(level) if s['name'] == sname), {})
    unit = sh.unit or cat.get('unit') or ''
    for k, u in ROW_UNITS.items():
        if k in chosen[2]:
            unit = u + ' (요약 행)'
    col_units = {c: u for _, c, n in sh.cols for k, u in ROW_UNITS.items() if k in n}
    return emit({'rows': choices, 'selected': chosen[1], 'name': chosen[2], 'unit': unit,
                 'meta': sh.meta, 'desc': cat.get('desc', ''), 'cells': cells,
                 'colUnits': col_units})


def search(payload):
    """모든 준비된 분류에서 산업명/코드 검색 + 상위 분류 경로."""
    q = json.loads(payload).get('q', '').strip().replace(' ', '')
    out = []
    if not q:
        return emit(out)
    for lv in LEVEL_ORDER:
        if lv not in DATA:
            continue
        m = DATA[lv]
        for i, c in enumerate(m.codes):
            if q in m.names[i].replace(' ', '') or c.startswith(q):
                path = [{'level': LEVELS[k]['label'], 'code': v, 'name': group_name(k, v)}
                        for k, v in parent_codes(lv, c).items() if v and k != lv]
                out.append({'level': lv, 'levelLabel': LEVELS[lv]['label'], 'code': c,
                            'name': m.names[i], 'x': m.x[i], 'path': path})
                if len(out) >= 300:
                    return emit(out)
    return emit(out)


def children(payload):
    """선택한 상위 분류 코드에 속하는 하위 분류 목록."""
    p = json.loads(payload)
    parent_level, code, child_level = p['parentLevel'], p['code'], p['childLevel']
    m = get_level(child_level)
    rows = []
    for i, c in enumerate(m.codes):
        if parent_codes(child_level, c).get(parent_level) == code:
            rows.append({'code': c, 'name': m.names[i], 'x': m.x[i],
                         'prodMult': m.L[:, i].sum(), 'vaMult': m.V[:, i].sum(),
                         'impMult': m.M[:, i].sum(), 'fd': m.fd[i]})
    return emit(rows)


# --- 산업구조 ---------------------------------------------------------------
def economy_indicators(m):
    x = m.x.sum()
    Zd, Zm = m.Zd.sum(), m.Zm.sum()
    imports = m.imports.sum()
    fd, fm = m.Fd.sum(), m.Fm.sum()
    exp = m.Fd[:, 6].sum() + m.Fm[:, 6].sum()
    P = m.L @ m.Fd
    Vv = m.V @ m.Fd
    cons = [k for k, (c, _) in enumerate(ITEMS) if ITEM_GROUP[c] == '소비']
    return {
        'x_total': x / 1e6, 'supply_total': (x + imports) / 1e6, 'import_total': imports / 1e6,
        'va_rate': m.va.sum() / x * 100, 'mid_rate': (Zd + Zm) / x * 100,
        'dom_mid_rate': Zd / x * 100, 'imp_dep': Zm / x * 100,
        'localization': Zd / (Zd + Zm) * 100, 'export_rate': exp / x * 100,
        'import_rate': imports / (x + imports) * 100,
        'prod_mult': P.sum() / fd, 'va_mult': Vv.sum() / fd,
        'imp_mult': (m.M @ m.fd).sum() / fd,
        'imp_mult_fd': ((m.M @ m.Fd).sum() + fm) / (fd + fm),
        'share_cons_prod': P[:, cons].sum() / P.sum() * 100,
        'share_exp_prod': P[:, 6].sum() / P.sum() * 100,
        'share_cons_va': Vv[:, cons].sum() / Vv.sum() * 100,
        'share_exp_va': Vv[:, 6].sum() / Vv.sum() * 100,
    }


def overview(level):
    m = get_level(level)
    x = m.x
    mid = m.Zd.sum(0) + m.Zm.sum(0)
    supply = x + m.imports
    exp = m.Fd[:, 6] + m.Fm[:, 6]
    rows = []
    for i, c in enumerate(m.codes):
        rows.append({
            'code': c, 'name': m.names[i], 'x': x[i], 'xShare': x[i] / x.sum() * 100,
            'va': m.va[i], 'vaRate': m.v[i] * 100,
            'midRate': safe_div(mid[i], x[i]) * 100,
            'impDep': safe_div(m.Zm[:, i].sum(), x[i]) * 100,
            'localization': safe_div(m.Zd[:, i].sum(), mid[i]) * 100,
            'wageShare': safe_div(m.VA[0, i], m.va[i]) * 100,
            'profitShare': safe_div(m.VA[1, i], m.va[i]) * 100,
            'export': exp[i], 'exportRate': safe_div(exp[i], x[i]) * 100,
            'imports': m.imports[i], 'importRate': safe_div(m.imports[i], supply[i]) * 100,
            'prodMult': m.L[:, i].sum(), 'vaMult': m.V[:, i].sum(), 'impMult': m.M[:, i].sum(),
            'active': bool(m.active[i])})
    eco = economy_indicators(m)
    bench = [{'label': lab, 'key': k, 'bok': val, 'calc': eco[k], 'unit': u,
              'diff': eco[k] - val} for lab, k, val, u in BOK['economy']]
    return emit({'rows': rows, 'economy': eco, 'bench': bench,
                 'label': LEVELS[level]['label'], 'source': m.source})


def bok_sector_check(_=None):
    out = []
    for lv in ('s7', 's13'):
        if lv not in DATA:
            continue
        m = DATA[lv]
        for c, (p, v, im) in BOK[lv].items():
            if c in m.idx:
                i = m.idx[c]
                out.append({'level': LEVELS[lv]['label'], 'code': c, 'name': m.names[i],
                            'bok': [p, v, im],
                            'calc': [m.L[:, i].sum(), m.V[:, i].sum(), m.M[:, i].sum()]})
    return emit(out)


# --- 수요충격 시뮬레이션 ----------------------------------------------------
def shock(payload):
    p = json.loads(payload)
    m = get_level(p['level'])
    f = np.zeros(m.n)
    for s in p['shocks']:
        a = float(s['amount'])
        if not np.isfinite(a):
            raise ValueError('수요 변화액은 유한한 수여야 합니다.')
        if s['code'] not in m.idx:
            raise ValueError(f'{s["code"]}: 이 분류에 없는 코드입니다.')
        f[m.idx[s['code']]] += a
    if not np.any(f):
        raise ValueError('0이 아닌 수요 변화액을 한 개 이상 입력하세요.')
    x = m.L @ f
    first = m.Ad @ f
    imp = m.Am @ x
    va = m.v * x
    wage = m.w * x
    profit = safe_div(m.VA[1], m.x) * x
    rows = [{'code': c, 'name': m.names[i], 'direct': f[i], 'first': first[i],
             'indirect': x[i] - f[i], 'production': x[i], 'imports': imp[i],
             'valueAdded': va[i], 'wage': wage[i]} for i, c in enumerate(m.codes)]
    groups = {}
    for target in ('large', 's13'):
        if LEVEL_ORDER.index(target) <= LEVEL_ORDER.index(m.level):
            continue
        agg = {}
        for i, c in enumerate(m.codes):
            k = parent_codes(m.level, c).get(target)
            if k is None:
                continue
            g = agg.setdefault(k, {'code': k, 'name': group_name(target, k),
                                   'production': 0.0, 'valueAdded': 0.0, 'imports': 0.0})
            g['production'] += x[i]
            g['valueAdded'] += va[i]
            g['imports'] += imp[i]
        groups[target] = sorted(agg.values(), key=lambda r: -abs(r['production']))
    tf = f.sum()
    same_sign = bool(np.all(f >= 0) or np.all(f <= 0))
    mult = (lambda a: a / tf) if (same_sign and tf) else (lambda a: None)
    return emit({
        'level': m.level, 'demand': tf, 'production': x.sum(), 'direct': tf,
        'firstRound': first.sum(), 'indirect': x.sum() - tf, 'imports': imp.sum(),
        'valueAdded': va.sum(), 'wage': wage.sum(), 'profit': profit.sum(),
        'prodMult': mult(x.sum()), 'vaMult': mult(va.sum()), 'impMult': mult(imp.sum()),
        'mixedSign': not same_sign,
        'check': abs(va.sum() + imp.sum() - tf) / max(abs(tf), 1),
        'rows': sorted(rows, key=lambda r: -abs(r['production'])), 'groups': groups})


# --- 전후방 연관효과 --------------------------------------------------------
def linkage(level):
    m = get_level(level)
    col = m.L.sum(0)
    row = m.L.sum(1)
    bl = col / col.mean()
    fl = row / row.mean()
    rows = []
    for i, c in enumerate(m.codes):
        q = ('Ⅰ' if bl[i] >= 1 and fl[i] >= 1 else 'Ⅱ' if bl[i] >= 1 else
             'Ⅲ' if fl[i] >= 1 else 'Ⅳ')
        rows.append({'code': c, 'name': m.names[i], 'colSum': col[i], 'rowSum': row[i],
                     'bl': bl[i], 'fl': fl[i], 'quad': q, 'x': m.x[i],
                     'vaMult': m.V[:, i].sum(), 'impMult': m.M[:, i].sum(),
                     'active': bool(m.active[i])})
    check = {}
    if level in BOOKS and 'L' in BOOKS[level]['roles']:
        sh = Sheet(BOOKS[level]['roles']['L'], BOOKS[level]['sheets'][BOOKS[level]['roles']['L']])
        r = find_code(sh, ['9801'], 'row', ('영향력',))
        c = find_code(sh, ['9802'], 'col', ('감응도',))
        if r:
            check['영향력계수 (계산 vs 원표)'] = float(np.max(np.abs(sh.matrix([r], m.codes)[0] - bl)))
        if c:
            check['감응도계수 (계산 vs 원표)'] = float(np.max(np.abs(sh.matrix(m.codes, [c])[:, 0] - fl)))
    counts = {q: sum(1 for r in rows if r['quad'] == q and r['active']) for q in 'ⅠⅡⅢⅣ'}
    return emit({'rows': rows, 'check': check, 'counts': counts})


# --- 최종수요 항목별 유발구조 ----------------------------------------------
def fd_structure(level):
    m = get_level(level)
    P = m.L @ m.Fd
    Vv = m.V @ m.Fd
    Mm = m.M @ m.Fd + m.Fm
    items = []
    for k, (c, nm) in enumerate(ITEMS):
        fd, fmk = m.Fd[:, k].sum(), m.Fm[:, k].sum()
        items.append({'code': c, 'name': nm, 'group': ITEM_GROUP[c], 'fd': fd, 'fm': fmk,
                      'prod': P[:, k].sum(), 'va': Vv[:, k].sum(), 'imp': Mm[:, k].sum(),
                      'prodMult': safe_div(P[:, k].sum(), fd),
                      'vaMult': safe_div(Vv[:, k].sum(), fd),
                      'impMult': safe_div(Mm[:, k].sum(), fd + fmk),
                      'prodShare': P[:, k].sum() / P.sum() * 100,
                      'vaShare': Vv[:, k].sum() / Vv.sum() * 100,
                      'impShare': Mm[:, k].sum() / Mm.sum() * 100})
    grp = {}
    for it in items:
        g = grp.setdefault(it['group'], {'group': it['group'], 'fd': 0, 'fm': 0, 'prod': 0,
                                         'va': 0, 'imp': 0})
        for key in ('fd', 'fm', 'prod', 'va', 'imp'):
            g[key] += it[key]
    for g in grp.values():
        g['prodMult'] = g['prod'] / g['fd']
        g['vaMult'] = g['va'] / g['fd']
        g['impMult'] = g['imp'] / (g['fd'] + g['fm'])
        g['prodShare'] = g['prod'] / P.sum() * 100
        g['vaShare'] = g['va'] / Vv.sum() * 100
        g['impShare'] = g['imp'] / Mm.sum() * 100
    gi = {'소비': [0, 1], '투자': [2, 3, 4, 5], '수출': [6]}
    rows = []
    xs = np.where(m.active, m.x, 1)
    for i, c in enumerate(m.codes):
        rows.append({'code': c, 'name': m.names[i], 'x': m.x[i],
                     'dep': {g: P[i, ks].sum() / xs[i] * 100 for g, ks in gi.items()},
                     'byItem': [P[i, k] for k in range(len(ITEMS))]})
    for it in items:
        it['bok'] = BOK['fd'].get(it['code'])
    for g in grp.values():
        g['bok'] = BOK['fd'].get(g['group'])
    return emit({'items': items, 'groups': [grp[g] for g in ('소비', '투자', '수출')],
                 'rows': rows, 'bokLevel': level == 'large'})


# --- 가격 파급효과(비용 인상의 물가 전가) -----------------------------------
def price(payload):
    """
    고정계수·완전전가 가정의 레온티에프 가격모형.
      p' = p' Ad + pm' Am + v'   →   Δp' = (Δpm' Am + Δv') (I - Ad)^-1
    - sector: 일부 부문의 가격을 외생적으로 r% 인상 → 나머지 부문으로 파급
    - import: 수입품 가격 r% 상승(전체 또는 선택 상품)
    - wage:   임금(피용자보수) r% 상승
    """
    p = json.loads(payload)
    m = get_level(p['level'])
    kind = p['kind']
    dp = np.zeros(m.n)
    dpm = np.zeros(m.n)
    exo = np.zeros(m.n, dtype=bool)
    if kind == 'sector':
        for s in p['targets']:
            i = m.idx[s['code']]
            exo[i] = True
            dp[i] = float(s['pct']) / 100
        if not exo.any():
            raise ValueError('가격을 올릴 부문을 한 개 이상 선택하세요.')
        N = ~exo
        AdN = m.Ad[np.ix_(N, N)]
        AdEN = m.Ad[np.ix_(exo, N)]
        dpN = (dp[exo] @ AdEN) @ np.linalg.inv(np.eye(N.sum()) - AdN)
        dp[N] = dpN
    elif kind == 'import':
        r = float(p['pct']) / 100
        targets = p.get('targets') or []
        if targets:
            for s in targets:
                dpm[m.idx[s['code']]] = r
        else:
            dpm[:] = r
        dp = (dpm @ m.Am) @ m.L
    elif kind == 'wage':
        r = float(p['pct']) / 100
        dp = (r * m.w) @ m.L
    else:
        raise ValueError('알 수 없는 분석 유형입니다.')
    cd, cm = m.Fd[:, 0], m.Fm[:, 0]
    exd = m.Fd[:, 6]
    idx = {
        'output': float(m.x @ dp / m.x.sum() * 100),
        'consumer': float((cd @ dp + cm @ dpm) / (cd.sum() + cm.sum()) * 100),
        'export': float(exd @ dp / exd.sum() * 100),
        'investment': float(((m.Fd[:, 2] + m.Fd[:, 3]) @ dp + (m.Fm[:, 2] + m.Fm[:, 3]) @ dpm) /
                            (m.Fd[:, 2:4].sum() + m.Fm[:, 2:4].sum()) * 100),
    }
    rows = [{'code': c, 'name': m.names[i], 'dp': dp[i] * 100, 'exo': bool(exo[i]),
             'x': m.x[i], 'consW': (cd[i] + cm[i]) / (cd.sum() + cm.sum()) * 100,
             'consContrib': (cd[i] * dp[i] + cm[i] * dpm[i]) / (cd.sum() + cm.sum()) * 100}
            for i, c in enumerate(m.codes)]
    rows.sort(key=lambda r: -abs(r['dp']))
    return emit({'kind': kind, 'index': idx, 'rows': rows})


# --- 분류 수준 비교 --------------------------------------------------------
def compare(payload):
    """
    (1) 수준별 요약: 부문 수, 경제 전체 유발계수(수준과 무관하게 같아야 함), 유발계수 분산
    (2) 공식 분류표 vs 기본부문 합산의 일치도
    (3) 같은 금액의 수요충격을 상위 분류 1개에 줄 때와, 그 하위 기본부문 중
        특정 부문에 줄 때(또는 실제 최종수요 구성비로 나눌 때)의 생산유발 차이(집계 편의)
    """
    p = json.loads(payload)
    summary = []
    for lv in LEVEL_ORDER:
        if lv not in DATA:
            continue
        m = DATA[lv]
        cs = m.L.sum(0)[m.active]
        e = economy_indicators(m)
        summary.append({'level': lv, 'label': LEVELS[lv]['label'], 'n': m.n, 'source': m.source,
                        'prodMult': e['prod_mult'], 'vaMult': e['va_mult'],
                        'impMult': e['imp_mult'], 'min': cs.min(), 'max': cs.max(),
                        'sd': cs.std(), 'x': m.x.sum()})
    consistency = []
    if 'basic' in DATA:
        for lv in ('small', 'medium', 'large'):
            if lv in DATA and DATA[lv].source in ('official', 'generated'):
                off = DATA[lv]
                agg = aggregate_from('basic', lv)
                common = [c for c in off.codes if c in agg.idx]
                if not common:
                    continue
                a = [agg.idx[c] for c in common]
                o = [off.idx[c] for c in common]
                consistency.append({
                    'level': LEVELS[lv]['label'], 'source': off.source, 'matched': len(common), 'n': off.n,
                    'xMaxRel': float(np.max(safe_div(np.abs(off.x[o] - agg.x[a]), np.abs(off.x[o])))),
                    'multMaxAbs': float(np.max(np.abs(off.L.sum(0)[o] - agg.L.sum(0)[a])))})
    detail = None
    if p.get('parentLevel') and p.get('code') and p.get('childLevel'):
        pl, code, cl = p['parentLevel'], p['code'], p['childLevel']
        amount = float(p.get('amount') or 1_000_000)
        P = get_level(pl)
        C = get_level(cl)
        if code not in P.idx:
            raise ValueError('상위 분류 코드가 없습니다.')
        i = P.idx[code]
        members = [j for j, c in enumerate(C.codes) if parent_codes(cl, c).get(pl) == code]
        if not members:
            raise ValueError('하위 분류 대응 부문이 없습니다(분류 수준 순서를 확인하세요).')
        fd = C.fd[members]
        wts = fd / fd.sum() if fd.sum() > 0 else np.ones(len(members)) / len(members)
        f = np.zeros(C.n)
        f[members] = amount * wts
        mix = (C.L @ f).sum()
        rows = []
        for j, w in zip(members, wts):
            rows.append({'code': C.codes[j], 'name': C.names[j], 'weight': w * 100,
                         'prodMult': C.L[:, j].sum(), 'vaMult': C.V[:, j].sum(),
                         'impMult': C.M[:, j].sum(), 'production': C.L[:, j].sum() * amount})
        rows.sort(key=lambda r: -r['weight'])
        detail = {'parent': {'code': code, 'name': P.names[i], 'label': LEVELS[pl]['label'],
                             'prodMult': P.L[:, i].sum(), 'vaMult': P.V[:, i].sum(),
                             'impMult': P.M[:, i].sum(), 'production': P.L[:, i].sum() * amount},
                  'childLabel': LEVELS[cl]['label'], 'amount': amount,
                  'mixProduction': mix, 'mixMult': mix / amount, 'rows': rows}
    return emit({'summary': summary, 'consistency': consistency, 'detail': detail})
