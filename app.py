# -*- coding: utf-8 -*-
import os
from flask import Flask, request, jsonify, render_template_string
import pandas as pd
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ──────────────────────────────────────────────
# 1. 배정예산 데이터 (백만원)
# ──────────────────────────────────────────────
BUDGET_CAPITAL = {
    #  사업명: (자금운용사업코드, 배정예산(백만원))
    '가공저압설비보강': ('310400201180203', 225),
    '취약선로보강': ('310400203200089', 199),
    '가공고압설비보강': ('310400208200094', 833),
    '노후 가공변압기 교체': ('310400212200088', 138),
    '과부하 가공변압기 교체': ('310400214200084', 89),
    '강도부족전주 교체': ('310400225200084', 182),
    '가공개폐기 교체': ('310400226200081', 88),
    'PCBS근절 변압기 교체': ('70H120224128', 180),
    '차량충돌복구공사': ('310400218200086', 41),
    '고장변압기 교체': ('310400213200087', 91),
    '재해설비 피해복구': ('70H120234279', 35),
    '가공설비고장복구': ('310400219200086', 115),
    '수선비/자본연계(가공보강)': ('S20245-575N-2019', 296),
}

CAPITAL_CAT_MAP = {
    '가공저압설비보강': '가공저압설비보강',
    '가공고압설비보강': '가공고압설비보강',
    '취약선로보강': '취약선로보강',
    '노후가공변압기보강': '노후 가공변압기 교체',
    '과부하가공변압기보강': '과부하 가공변압기 교체',
    '강도부족전주보강': '강도부족전주 교체',
    '가공개폐기보강': '가공개폐기 교체',
    'PCBS근절 변압기 교체': 'PCBS근절 변압기 교체',
    '차량충돌복구공사': '차량충돌복구공사',
    '고장변압기복구': '고장변압기 교체',
    '재해설비 피해복구': '재해설비 피해복구',
    '가공설비고장복구': '가공설비고장복구',
}

# ── 손익(수선비) 배정예산 ──
BUDGET_REVENUE = {
    '배전경상': ('S10030', 37),
    '내선경상': ('S10040', 2),
    '가공배전설비진단': ('S20110', 73),
    '배전수목전지': ('S20140', 91),
    '배전조류고장예방': ('S20150', 76),
    '배전염진해낙뢰': ('S20160', 45),
    '가공배전기타계획(주요설비)': ('S20176', 105),
    '가공배전기타계획(기타설비)': ('S20180', 79),
    '계기함정비': ('S20190', 13),
    '자본연계(공급_선진)': ('S20240', 18),
    '자본연계(보강/가공)': ('S20245', 796),
    '지중배전기타점검(주요설비)': ('S20265', 8),
    '배전기자재수리': ('S20290', 0),
    '자본연계(신규)': ('S20220', 131),
    '자본연계(지장)': ('S20230', 525),
    '자본연계(보강/내선)': ('S20247', 141),
}

REVENUE_CAT_MAP = {
    '배전경상': '배전경상',
    '배전수목전지': '배전수목전지',
    '배전조류고장예방': '배전조류고장예방',
    '가공배전기타계획': '가공배전기타계획(주요설비)',
    '배전접지보강': '배전염진해낙뢰',
    '배전변압기개폐기': '가공배전설비진단',
    '도서S/C관리': '가공배전기타계획(기타설비)',
    '자본연계(가공보강)': '자본연계(보강/가공)',
    '자본연계(신규)': '자본연계(신규)',
    '자본연계(지장)': '자본연계(지장)',
    '자본연(도서보강배전)': '자본연계(공급_선진)',
    '자본연계(도서배전)': '자본연계(공급_선진)',
    '자본연계(배전계획)': '계기함정비',
    '재해손실(기타)': '배전기자재수리',
    '자본연계(지중보강)': '자본연계(보강/내선)',
}


# ──────────────────────────────────────────────
# 2. 유틸리티
# ──────────────────────────────────────────────
def _num(val):
    if pd.isna(val):
        return 0
    if isinstance(val, str):
        val = val.replace(',', '').strip()
    try:
        return round(float(val))
    except (ValueError, TypeError):
        return 0


def _pct(val):
    try:
        v = float(val)
        if pd.isna(v):
            return 0.0
        return round(v * 100, 1) if v <= 1.5 else round(v, 1)
    except (ValueError, TypeError):
        return 0.0


def _date_str(val):
    if pd.isna(val):
        return ''
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    return str(val).strip()[:10]


# ──────────────────────────────────────────────
# 3. 파싱 + 분석
# ──────────────────────────────────────────────
def parse_and_analyze(filepath):
    """엑셀 파일을 파싱하여 자본/손익 분리 분석 결과를 반환한다."""
    xls = pd.ExcelFile(filepath)

    # ── 공사관리대장 시트 찾기 ──
    ledger_sheet = None
    for name in xls.sheet_names:
        if '공사관리' in name or '세부' in name:
            ledger_sheet = name
            break
    if ledger_sheet is None and len(xls.sheet_names) > 0:
        ledger_sheet = xls.sheet_names[0]

    df = pd.read_excel(xls, sheet_name=ledger_sheet, header=None)
    header_row = 1
    for i in range(min(5, len(df))):
        if pd.notna(df.iloc[i, 0]) and '순번' in str(df.iloc[i, 0]):
            header_row = i
            break

    projects = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        cno = row[1] if 1 < df.shape[1] else None
        if pd.isna(cno) or str(cno).strip() == '':
            continue

        cat_capital = str(row[13]).strip() if 13 < df.shape[1] and pd.notna(row.get(13)) else '미분류'
        # 손익예산과목: "배전경상-경남/통영/전력공급팀" → "배전경상"
        raw_rev = str(row[14]).strip() if 14 < df.shape[1] and pd.notna(row.get(14)) else '미분류'
        cat_revenue = raw_rev.split('-')[0].strip() if '-' in raw_rev else raw_rev
        status = str(row[26]).strip() if 26 < df.shape[1] and pd.notna(row.get(26)) else '미확인'

        projects.append({
            '공사번호': str(cno).strip(),
            '공사성격': str(row[4]).strip() if 4 < df.shape[1] and pd.notna(row.get(4)) else '',
            '공사업체': str(row[6]).strip() if 6 < df.shape[1] and pd.notna(row.get(6)) else '',
            '자본예산과목': cat_capital,
            '손익예산과목': cat_revenue,
            '총공사비_설계': _num(row[15]) if 15 < df.shape[1] else 0,
            '도급비_계약': _num(row[17]) if 17 < df.shape[1] else 0,
            '착공일': _date_str(row[18]) if 18 < df.shape[1] else '',
            '공사상태': status,
            '설계_자본': _num(row[41]) if 41 < df.shape[1] else 0,
            '기성_자본': _num(row[47]) if 47 < df.shape[1] else 0,
            '예정_자본': _num(row[51]) if 51 < df.shape[1] else 0,
            '설계_손익': _num(row[42]) if 42 < df.shape[1] else 0,
            '기성_손익': _num(row[48]) if 48 < df.shape[1] else 0,
            '예정_손익': _num(row[52]) if 52 < df.shape[1] else 0,
        })

    # ── 기성고 시트 ──
    inspections = []
    for name in xls.sheet_names:
        if '기성고' in name:
            df2 = pd.read_excel(xls, sheet_name=name, header=None)
            hr = 1
            for i in range(min(5, len(df2))):
                if pd.notna(df2.iloc[i, 0]) and '순번' in str(df2.iloc[i, 0]):
                    hr = i
                    break
            for i in range(hr + 1, len(df2)):
                row = df2.iloc[i]
                cno = row[1] if 1 < df2.shape[1] else None
                if pd.isna(cno):
                    continue
                inspections.append({
                    '공사번호': str(cno).strip(),
                    '기성금액': _num(row[7]) if 7 < df2.shape[1] else 0,
                    '기성검사보고일': _date_str(row[6]) if 6 < df2.shape[1] else '',
                })
            break

    # ── 예산현황 시트 파싱 ──
    budget_sheets = {}
    for sheet_key, sheet_name, skip_kw in [
        ('전기품질', '전기품질 자본', ['소계']),
        ('지장주', '지장주 자본', ['소계']),
        ('수선비', '수선비', ['소계', '합계']),
    ]:
        if sheet_name not in xls.sheet_names:
            continue
        sdf = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        rows = []
        for i in range(2, len(sdf)):
            r = sdf.iloc[i]
            nm = r[1]
            if pd.isna(nm) or any(kw in str(nm) for kw in skip_kw):
                continue
            rows.append({
                '사업코드': str(r[0]).strip() if pd.notna(r[0]) else '',
                '사업명': str(nm).strip(),
                '최초배정예산': _num(r[2]),
                '배정예산': _num(r[3]),
                '소비금액': _num(r[4]),
                '약정금액': _num(r[5]),
                '집행실적': _num(r[6]),
                '잔액': _num(r[7]),
                '집행율': _pct(r[8]),
                '진행중공사비': _num(r[9]),
                '최종예상집행실적': _num(r[10]),
                '예상잔액': _num(r[11]),
                '최종예상집행율': _pct(r[12]),
            })
        budget_sheets[sheet_key] = rows

    # ═══════════════════════════════════════
    # 분석: 자본 / 손익 분리
    # ═══════════════════════════════════════

    def _aggregate(proj_list, design_key, paid_key, est_key, cat_key,
                   budget_dict=None, cat_map_dict=None):
        """과목별 집계 (엑셀 예산현황 시트와 동일 구조)"""
        cat_map = {}
        for p in proj_list:
            cat = p[cat_key]
            if cat not in cat_map:
                cat_map[cat] = {
                    '과목': cat, '건수': 0, '완료': 0, '진행중': 0, '중지': 0,
                    '소비금액': 0, '약정금액': 0, '진행중공사비': 0,
                }
            s = cat_map[cat]
            s['건수'] += 1
            if '완료' in p['공사상태']:
                s['완료'] += 1
            elif '중지' in p['공사상태']:
                s['중지'] += 1
                diff = p[design_key] - p[paid_key]
                if diff > 0:
                    s['약정금액'] += diff
            else:
                s['진행중'] += 1
                diff = p[design_key] - p[paid_key]
                if diff > 0:
                    s['약정금액'] += diff
            s['소비금액'] += p[paid_key]
            s['진행중공사비'] += p[est_key]

        # 사업코드 매핑
        reverse_map = {}
        if cat_map_dict and budget_dict:
            for cat_name, budget_name in cat_map_dict.items():
                if budget_name in budget_dict:
                    reverse_map[cat_name] = budget_dict[budget_name][0]

        for s in cat_map.values():
            s['사업코드'] = reverse_map.get(s['과목'], '')
            b = s['소비금액']
            c = s['약정금액']
            s['집행실적'] = b + c
            f = s['진행중공사비']
            s['예상집행'] = s['집행실적'] + f

        return sorted(cat_map.values(), key=lambda x: x['소비금액'] + x['약정금액'], reverse=True)

    cap_cats = _aggregate(projects, '설계_자본', '기성_자본', '예정_자본', '자본예산과목',
                          BUDGET_CAPITAL, CAPITAL_CAT_MAP)
    rev_cats = _aggregate(projects, '설계_손익', '기성_손익', '예정_손익', '손익예산과목',
                          BUDGET_REVENUE, REVENUE_CAT_MAP)

    # ── 자본 배정예산 대비 ──
    def _build_comparison(budget_dict, cat_map_dict, cat_list):
        cat_dict = {c['과목']: c for c in cat_list}
        comparison = []
        for bname, (bcode, bval_mil) in budget_dict.items():
            bval = bval_mil * 1_000_000
            matched = [cat for cat, mapped in cat_map_dict.items() if mapped == bname]
            sobi = sum(cat_dict.get(c, {}).get('소비금액', 0) for c in matched)
            yakjung = sum(cat_dict.get(c, {}).get('약정금액', 0) for c in matched)
            exec_total = sobi + yakjung
            progress = sum(cat_dict.get(c, {}).get('진행중공사비', 0) for c in matched)
            forecast = exec_total + progress
            count = sum(cat_dict.get(c, {}).get('건수', 0) for c in matched)
            exec_rate = round(exec_total / bval * 100, 1) if bval else 0
            fc_rate = round(forecast / bval * 100, 1) if bval else 0
            comparison.append({
                '사업코드': bcode, '예산과목': bname,
                '배정예산_백만': bval_mil, '배정예산': bval,
                '소비금액': sobi, '약정금액': yakjung,
                '집행실적': exec_total, '잔액': bval - exec_total,
                '집행율': exec_rate,
                '진행중공사비': progress, '예상집행': forecast,
                '예상잔액': bval - forecast, '예상집행율': fc_rate,
                '건수': count,
                '상태': '초과' if fc_rate > 100 else ('양호' if fc_rate > 70 else '미달'),
            })
        return comparison

    cap_comparison = _build_comparison(BUDGET_CAPITAL, CAPITAL_CAT_MAP, cap_cats)
    rev_comparison = _build_comparison(BUDGET_REVENUE, REVENUE_CAT_MAP, rev_cats)

    # 공사상태별
    def _status_agg(proj_list, design_key):
        sm = {}
        for p in proj_list:
            st = p['공사상태']
            if st not in sm:
                sm[st] = {'건수': 0, '금액': 0}
            sm[st]['건수'] += 1
            sm[st]['금액'] += p[design_key]
        return sm

    cap_status = _status_agg(projects, '설계_자본')
    rev_status = _status_agg(projects, '설계_손익')

    # 총계 계산 함수
    def _totals(comparison, budget_dict):
        total_budget = sum(v[1] * 1e6 for v in budget_dict.values())
        total_sobi = sum(c['소비금액'] for c in comparison)
        total_yakjung = sum(c['약정금액'] for c in comparison)
        total_exec = total_sobi + total_yakjung
        total_progress = sum(c['진행중공사비'] for c in comparison)
        total_forecast = total_exec + total_progress
        total_count = sum(c['건수'] for c in comparison)
        return {
            '배정예산': total_budget,
            '소비금액': total_sobi,
            '약정금액': total_yakjung,
            '집행실적': total_exec,
            '잔액': total_budget - total_exec,
            '집행율': round(total_exec / total_budget * 100, 1) if total_budget else 0,
            '진행중공사비': total_progress,
            '예상집행': total_forecast,
            '예상잔액': total_budget - total_forecast,
            '예상집행율': round(total_forecast / total_budget * 100, 1) if total_budget else 0,
            '공사건수': total_count,
            '초과항목': sum(1 for m in comparison if m['상태'] == '초과'),
        }

    cap_summary = _totals(cap_comparison, BUDGET_CAPITAL)
    rev_summary = _totals(rev_comparison, BUDGET_REVENUE)

    return {
        'capital': {
            'summary': cap_summary,
            'budget_comparison': cap_comparison,
            'category': cap_cats,
            'status': {k: v for k, v in cap_status.items()},
            'budget_sheets': {k: v for k, v in budget_sheets.items() if k in ('전기품질', '지장주')},
        },
        'revenue': {
            'summary': rev_summary,
            'budget_comparison': rev_comparison,
            'category': rev_cats,
            'status': {k: v for k, v in rev_status.items()},
            'budget_sheets': {k: v for k, v in budget_sheets.items() if k == '수선비'},
        },
        'projects': projects[:500],
    }


# ──────────────────────────────────────────────
# 4. HTML 템플릿
# ──────────────────────────────────────────────
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>배전공사 예산 관리</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{
  --navy:#0D2137;--navy2:#1A3A5C;--navy3:#2C5282;--navy-light:#3B6AA0;
  --orange:#E8731B;--orange-light:#F5913D;--orange-bg:#FFF4EB;
  --green:#10B981;--green-bg:#D1FAE5;--amber:#F59E0B;--amber-bg:#FEF3C7;
  --red:#EF4444;--red-bg:#FEE2E2;--purple:#7C3AED;
  --bg:#F0F4F8;--card:#FFFFFF;--border:#E2E8F0;
  --text:#1A202C;--text2:#64748B;--text3:#94A3B8;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Malgun Gothic','Apple SD Gothic Neo','Segoe UI',sans-serif;background:var(--bg);color:var(--text);font-size:13px}

/* ── Header ── */
header{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 60%,var(--navy3) 100%);color:#fff;padding:0 28px;height:56px;display:flex;justify-content:space-between;align-items:center;box-shadow:0 2px 12px rgba(0,0,0,.2)}
header h1{font-size:18px;font-weight:700;letter-spacing:-.3px;display:flex;align-items:center;gap:8px}
header h1 .logo{width:28px;height:28px;background:var(--orange);border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:800;color:#fff}
header .rt{display:flex;align-items:center;gap:10px}
.ubtn{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.25);padding:6px 14px;border-radius:6px;font-weight:600;font-size:11px;cursor:pointer;transition:all .15s;backdrop-filter:blur(4px)}
.ubtn:hover{background:rgba(255,255,255,.25);border-color:rgba(255,255,255,.4)}
.ubtn-primary{background:var(--orange);border-color:var(--orange);color:#fff}
.ubtn-primary:hover{background:var(--orange-light)}
.sp{display:none}.sp.on{display:inline-block;width:14px;height:14px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── Layout ── */
.wrap{max-width:1600px;margin:0 auto;padding:20px 24px}
.main-tabs{display:flex;gap:2px;margin-bottom:20px}
.mt{padding:11px 36px;font-size:14px;font-weight:700;border:none;cursor:pointer;border-radius:8px 8px 0 0;transition:all .2s;letter-spacing:-.2px}
.mt.cap{background:var(--navy2);color:#fff}.mt.rev{background:var(--orange);color:#fff}
.mt.off{background:var(--border);color:var(--text3)}
.mt.off:hover{background:#d1d5db;color:var(--text2)}
.main-pane{display:none}.main-pane.on{display:block}

/* ── Sub tabs ── */
.stabs{display:flex;border-bottom:2px solid var(--border);margin-bottom:0}
.st{padding:9px 20px;background:transparent;border:none;font-size:12px;font-weight:600;color:var(--text3);cursor:pointer;border-bottom:2px solid transparent;transition:all .15s}
.st.on{color:var(--navy2);border-bottom-color:var(--navy2)}.st:hover{color:var(--navy2)}
.sp2{display:none}.sp2.on{display:block}

/* ── Cards ── */
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:20px}
.cd{background:var(--card);border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.05);position:relative;overflow:hidden;border:1px solid var(--border);transition:box-shadow .15s}
.cd:hover{box-shadow:0 3px 12px rgba(0,0,0,.08)}
.cd::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;border-radius:4px 0 0 4px}
.cd.c1::before{background:var(--navy2)}.cd.c2::before{background:var(--green)}.cd.c3::before{background:var(--amber)}
.cd.c4::before{background:var(--red)}.cd.c5::before{background:var(--purple)}.cd.c6::before{background:var(--navy-light)}
.cd .lb{font-size:10px;color:var(--text3);margin-bottom:4px;font-weight:600;text-transform:uppercase;letter-spacing:.3px}
.cd .vl{font-size:19px;font-weight:800;color:var(--text);letter-spacing:-.3px}
.cd .sb{font-size:10px;color:var(--text3);margin-top:3px}

/* ── Chart / Table boxes ── */
.cbox{background:var(--card);border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.05);margin:16px 0;border:1px solid var(--border)}
.cbox canvas{max-height:320px}
.cbox h3{font-size:13px;color:var(--text2);margin-bottom:10px;font-weight:700}
.tb{background:var(--card);border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.05);margin-top:14px;overflow-x:auto;border:1px solid var(--border)}
.tb h3{font-size:13px;color:var(--text2);margin-bottom:10px;font-weight:700}
.sc{max-height:520px;overflow-y:auto}

/* ── Table ── */
table{width:100%;border-collapse:collapse;font-size:11px}
thead th{background:#F8FAFC;padding:8px 3px;text-align:right;border-bottom:2px solid var(--border);white-space:nowrap;font-weight:700;color:var(--text2);position:sticky;top:0;z-index:1;font-size:10px}
thead th:first-child,thead th:nth-child(2){text-align:left}
tbody td{padding:7px 3px;border-bottom:1px solid #F1F5F9;text-align:right;white-space:nowrap}
tbody td:first-child,tbody td:nth-child(2){text-align:left;max-width:170px;overflow:hidden;text-overflow:ellipsis}
tbody tr:hover{background:#F8FAFC}
tbody tr:nth-child(even){background:#FDFDFE}
tfoot td{padding:8px 3px;font-weight:800;border-top:2px solid var(--navy2);text-align:right;background:#F0F4FF;color:var(--navy2)}
tfoot td:first-child,tfoot td:nth-child(2){text-align:left}

/* ── Progress bar ── */
.br{display:inline-block;width:52px;height:5px;background:#E2E8F0;border-radius:3px;overflow:hidden;vertical-align:middle;margin-right:3px}
.br .f{display:block;height:100%;border-radius:3px}
.br.lo .f{background:var(--red)}.br.md .f{background:var(--amber)}.br.hi .f{background:var(--green)}

/* ── Badges ── */
.bg{padding:2px 7px;border-radius:4px;font-size:9px;font-weight:700;letter-spacing:.2px}
.bg-ok{background:var(--green-bg);color:#065F46}.bg-wn{background:var(--amber-bg);color:#92400E}.bg-ov{background:var(--red-bg);color:#991B1B}
.bg-dn{background:#DBEAFE;color:#1E40AF}.bg-wp{background:var(--orange-bg);color:#9A3412}.bg-st{background:#F1F5F9;color:#475569}

/* ── Input ── */
.inp-budget{width:130px;padding:4px 6px;border:1px solid var(--border);border-radius:4px;text-align:right;font-size:11px;font-weight:700;background:#FFFBEB;transition:all .15s}
.inp-budget:focus{border-color:var(--navy2);outline:none;background:#fff;box-shadow:0 0 0 2px rgba(26,58,92,.15)}
.inp-code{width:120px;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:11px;transition:all .15s}
.inp-code:focus{border-color:var(--navy2);outline:none;box-shadow:0 0 0 2px rgba(26,58,92,.15)}

/* ── Search ── */
.sf{margin:12px 0;display:flex;gap:8px;align-items:center}
.sf input,.sf select{padding:7px 12px;border:1px solid var(--border);border-radius:6px;font-size:12px;transition:border-color .15s}
.sf input{width:280px}
.sf input:focus,.sf select:focus{border-color:var(--navy2);outline:none}

/* ── Colors ── */
.neg{color:var(--red);font-weight:600}.pos{color:var(--green);font-weight:600}

/* ── Add row button ── */
.add-row-btn{margin-top:10px;padding:7px 18px;background:var(--card);border:1px dashed var(--navy-light);color:var(--navy2);border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s}
.add-row-btn:hover{background:var(--navy2);color:#fff;border-style:solid}

/* ── Responsive ── */
@media(max-width:1200px){.cards{grid-template-columns:repeat(3,1fr)}}
@media(max-width:900px){.cards{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){header{flex-direction:column;height:auto;padding:12px 16px;gap:8px}.cards{grid-template-columns:1fr}.wrap{padding:12px}}
</style>
</head>
<body>
<header>
<h1>배전공사 예산 관리 시스템</h1>
<div class="rt">
<span class="sp" id="sp"></span>
<span id="fn" style="font-size:11px;opacity:.8"></span>
<span id="today" style="font-size:11px;opacity:.8"></span>
<button class="ubtn" id="refreshBtn" onclick="doRefresh()" style="display:none">&#8635; 새로고침</button>
<label class="ubtn ubtn-primary">파일 업로드<input type="file" id="fi" accept=".xlsx,.xls" style="display:none"></label>
</div>
</header>
<div class="wrap">
<div class="main-tabs">
<button class="mt cap" id="mtCap" onclick="switchMain('cap')">자본</button>
<button class="mt rev off" id="mtRev" onclick="switchMain('rev')">손익</button>
</div>

<!-- ═══ 자본 ═══ -->
<div class="main-pane on" id="mpCap">
<div class="cards">
<div class="cd c1"><div class="lb">배정예산 (A)</div><div class="vl" id="capA">-</div><div class="sb">자본 배정예산</div></div>
<div class="cd c2"><div class="lb">집행실적 (D=B+C)</div><div class="vl" id="capD">-</div><div class="sb" id="capDR">집행율 -</div></div>
<div class="cd c4"><div class="lb">잔액 (A-D)</div><div class="vl" id="capE">-</div><div class="sb">배정예산 - 집행실적</div></div>
<div class="cd c3"><div class="lb">진행중공사비 (F)</div><div class="vl" id="capF">-</div><div class="sb">미준공 금액</div></div>
<div class="cd c5"><div class="lb">최종예상 (G=D+F)</div><div class="vl" id="capG">-</div><div class="sb" id="capGR">예상집행율 -</div></div>
<div class="cd c6"><div class="lb">공사건수</div><div class="vl" id="capCnt">-</div><div class="sb" id="capOver">초과 -</div></div>
</div>
<div class="stabs">
<button class="st on" onclick="subTab('cap',this,'capBudget')">예산현황</button>
<button class="st" onclick="subTab('cap',this,'capProj')">공사목록</button>
</div>
<div class="sp2 on" id="capBudgetP">
<div class="cbox"><h3>자본 예산 집행 현황 (백만원)</h3><canvas id="chCapBar"></canvas></div>
<div class="tb"><h3>자본 예산현황 (단위: 원, 배정예산 직접 입력)</h3><div class="sc">
<table id="tCapMain">
<thead><tr>
<th>자금운용<br>사업코드</th><th>사업명</th><th>배정예산<br>(A)</th>
<th>소비금액<br>(B)</th><th>약정금액<br>(C)</th><th>집행실적<br>(D=B+C)</th>
<th>잔액<br>(E=A-D)</th><th>집행율<br>(D/A)</th>
<th>진행중<br>공사비(F)</th><th>최종예상<br>(G=D+F)</th><th>예상잔액<br>(A-G)</th><th>예상집행율<br>(G/A)</th><th>공사<br>건수</th>
</tr></thead>
<tbody></tbody>
<tfoot><tr id="capTotalRow">
<td></td><td>합 계</td><td id="capTotA">-</td><td id="capTotB">-</td><td id="capTotC">-</td>
<td id="capTotD">-</td><td id="capTotE">-</td><td id="capTotDR">-</td>
<td id="capTotF">-</td><td id="capTotG">-</td><td id="capTotGA">-</td><td id="capTotGR">-</td><td id="capTotN">-</td>
</tr></tfoot>
</table></div>
<button class="add-row-btn" onclick="addRow('cap')">+ 항목 추가</button>
</div>
</div>
<div class="sp2" id="capProjP">
<div class="sf"><input type="text" id="capSrch" placeholder="공사번호/업체명/과목 검색..."><select id="capFilt"><option value="">전체</option><option value="공사완료">완료</option><option value="공사중">진행</option><option value="공사중지">중지</option></select></div>
<div class="tb"><h3>공사목록 (자본)</h3><div class="sc"><table id="tCapProj"><thead><tr><th>No</th><th>공사번호</th><th>자본예산과목</th><th>공사업체</th><th>상태</th><th>착공일</th><th>설계(자본)</th><th>기성(자본)</th><th>예정(자본)</th><th>기성율</th></tr></thead><tbody></tbody></table></div></div>
</div>
</div>

<!-- ═══ 손익 ═══ -->
<div class="main-pane" id="mpRev">
<div class="cards">
<div class="cd c1"><div class="lb">배정예산 (A)</div><div class="vl" id="revA">-</div><div class="sb">손익 배정예산</div></div>
<div class="cd c2"><div class="lb">집행실적 (D=B+C)</div><div class="vl" id="revD">-</div><div class="sb" id="revDR">집행율 -</div></div>
<div class="cd c4"><div class="lb">잔액 (A-D)</div><div class="vl" id="revE">-</div><div class="sb">배정예산 - 집행실적</div></div>
<div class="cd c3"><div class="lb">진행중공사비 (F)</div><div class="vl" id="revF">-</div><div class="sb">미준공 금액</div></div>
<div class="cd c5"><div class="lb">최종예상 (G=D+F)</div><div class="vl" id="revG">-</div><div class="sb" id="revGR">예상집행율 -</div></div>
<div class="cd c6"><div class="lb">공사건수</div><div class="vl" id="revCnt">-</div><div class="sb" id="revOver">초과 -</div></div>
</div>
<div class="stabs">
<button class="st on" onclick="subTab('rev',this,'revBudget')">예산현황</button>
<button class="st" onclick="subTab('rev',this,'revProj')">공사목록</button>
</div>
<div class="sp2 on" id="revBudgetP">
<div class="cbox"><h3>손익 예산 집행 현황 (백만원)</h3><canvas id="chRevBar"></canvas></div>
<div class="tb"><h3>손익 예산현황 (단위: 원, 배정예산 직접 입력)</h3><div class="sc">
<table id="tRevMain">
<thead><tr>
<th>자금운용<br>사업코드</th><th>사업명</th><th>배정예산<br>(A)</th>
<th>소비금액<br>(B)</th><th>약정금액<br>(C)</th><th>집행실적<br>(D=B+C)</th>
<th>잔액<br>(E=A-D)</th><th>집행율<br>(D/A)</th>
<th>진행중<br>공사비(F)</th><th>최종예상<br>(G=D+F)</th><th>예상잔액<br>(A-G)</th><th>예상집행율<br>(G/A)</th><th>공사<br>건수</th>
</tr></thead>
<tbody></tbody>
<tfoot><tr id="revTotalRow">
<td></td><td>합 계</td><td id="revTotA">-</td><td id="revTotB">-</td><td id="revTotC">-</td>
<td id="revTotD">-</td><td id="revTotE">-</td><td id="revTotDR">-</td>
<td id="revTotF">-</td><td id="revTotG">-</td><td id="revTotGA">-</td><td id="revTotGR">-</td><td id="revTotN">-</td>
</tr></tfoot>
</table></div>
<button class="add-row-btn" onclick="addRow('rev')">+ 항목 추가</button>
</div>
</div>
<div class="sp2" id="revProjP">
<div class="sf"><input type="text" id="revSrch" placeholder="공사번호/업체명/과목 검색..."><select id="revFilt"><option value="">전체</option><option value="공사완료">완료</option><option value="공사중">진행</option><option value="공사중지">중지</option></select></div>
<div class="tb"><h3>공사목록 (손익)</h3><div class="sc"><table id="tRevProj"><thead><tr><th>No</th><th>공사번호</th><th>손익예산과목</th><th>공사업체</th><th>상태</th><th>착공일</th><th>설계(손익)</th><th>기성(손익)</th><th>예정(손익)</th><th>기성율</th></tr></thead><tbody></tbody></table></div></div>
</div>
</div>
</div>

<script>
let D=null,CH={};

function switchMain(t){
    document.querySelectorAll('.mt').forEach(b=>b.classList.add('off'));
    document.querySelectorAll('.main-pane').forEach(p=>p.classList.remove('on'));
    if(t==='cap'){document.getElementById('mtCap').classList.remove('off');document.getElementById('mpCap').classList.add('on')}
    else{document.getElementById('mtRev').classList.remove('off');document.getElementById('mpRev').classList.add('on')}
}

function subTab(prefix,btn,pane){
    btn.closest('.stabs').querySelectorAll('.st').forEach(b=>b.classList.remove('on'));btn.classList.add('on');
    btn.closest('.main-pane').querySelectorAll('.sp2').forEach(p=>p.classList.remove('on'));
    document.getElementById(pane+'P').classList.add('on');
}

document.getElementById('fi').addEventListener('change',async e=>{
    const f=e.target.files[0];if(!f)return;
    document.getElementById('fn').textContent=f.name;
    const sp=document.getElementById('sp');sp.classList.add('on');
    const fd=new FormData();fd.append('file',f);
    try{const r=await fetch('/api/analyze',{method:'POST',body:fd});const j=await r.json();if(j.error){alert(j.error);return}D=j;document.getElementById('refreshBtn').style.display='inline-block';renderAll()}
    catch(err){alert(err.message)}finally{sp.classList.remove('on')}
});

async function doRefresh(){
    const sp=document.getElementById('sp');sp.classList.add('on');
    try{const r=await fetch('/api/default');const j=await r.json();if(!j.error){D=j;renderAll()}}
    catch(e){}finally{sp.classList.remove('on')}
}

const fm=v=>(v/1e6).toLocaleString('ko-KR',{maximumFractionDigits:0})+' 백만';
const fw=v=>Math.round(v).toLocaleString('ko-KR');
const fp=v=>v.toFixed(1)+'%';
function br(r){const c=r>100?'neg':r>70?'pos':'';return `<span class="${c}">${fp(r)}</span>`}
function bg(s){if(s==='초과')return'<span class="bg bg-ov">초과</span>';if(s==='양호')return'<span class="bg bg-ok">양호</span>';return'<span class="bg bg-wn">미달</span>'}
function stBg(s){if(s.includes('완료'))return'<span class="bg bg-dn">완료</span>';if(s.includes('중지'))return'<span class="bg bg-st">중지</span>';return'<span class="bg bg-wp">진행</span>'}
function sn(n,l){return n.length>l?n.slice(0,l)+'..':n}
function clr(v){return v<0?'neg':'pos'}
function niceMax(v){if(v<=0)return 100;const t=v*1.15;if(t<=100)return Math.ceil(t/50)*50;return Math.ceil(t/200)*200}
function _ch(id,type,data,opts={}){if(CH[id])CH[id].destroy();CH[id]=new Chart(document.getElementById(id),{type,data,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:{size:11}}}},...opts}})}

function renderAll(){if(!D)return;renderCapital();renderRevenue();renderProjects()}

// ═══════════════════════════════════════
// 자본 렌더
// ═══════════════════════════════════════
function renderCapital(){
    const s=D.capital.summary;
    document.getElementById('capA').textContent=fm(s['배정예산']);
    document.getElementById('capD').textContent=fm(s['집행실적']);
    document.getElementById('capDR').textContent='집행율 '+fp(s['집행율']);
    document.getElementById('capE').textContent=fm(s['잔액']);
    document.getElementById('capF').textContent=fm(s['진행중공사비']);
    document.getElementById('capG').textContent=fm(s['예상집행']);
    document.getElementById('capGR').textContent='예상집행율 '+fp(s['예상집행율']);
    document.getElementById('capCnt').textContent=s['공사건수']+'건';
    document.getElementById('capOver').textContent='초과 '+s['초과항목']+'건';

    const comp=D.capital.budget_comparison;
    const lb=[],aV=[],dV=[],gV=[];
    comp.forEach(r=>{if(!r['배정예산_백만'])return;lb.push(sn(r['예산과목'],7));aV.push(r['배정예산_백만']);dV.push(Math.round(r['집행실적']/1e6));gV.push(Math.round(r['예상집행']/1e6))});
    const capMax=Math.max(...aV,...dV,...gV);const capYMax=niceMax(capMax);
    _ch('chCapBar','bar',{labels:lb,datasets:[
        {label:'배정예산(A)',data:aV,backgroundColor:'rgba(59,130,246,.7)',borderColor:'rgba(59,130,246,1)',borderWidth:1,borderRadius:3},
        {label:'집행실적(D)',data:dV,backgroundColor:'rgba(239,68,68,.65)',borderColor:'rgba(239,68,68,1)',borderWidth:1,borderRadius:3},
        {label:'최종예상(G)',data:gV,backgroundColor:'rgba(251,191,36,.6)',borderColor:'rgba(245,158,11,1)',borderWidth:1,borderRadius:3}
    ]},{scales:{y:{beginAtZero:true,max:capYMax,afterBuildTicks(axis){const t=[0,50,100];for(let v=200;v<=axis.max;v+=200)t.push(v);axis.ticks=t.map(v=>({value:v}))},ticks:{font:{size:10}},title:{display:true,text:'백만원'}}}});

    const tb=document.querySelector('#tCapMain tbody');tb.innerHTML='';
    comp.forEach((r,i)=>{
        const a=r['배정예산'],b=r['소비금액'],c=r['약정금액'],d=r['집행실적'],e=r['잔액'],f=r['진행중공사비'],g=r['예상집행'],ga=r['예상잔액'],dr=r['집행율'],gr=r['예상집행율'];
        const tr=document.createElement('tr');
        tr.innerHTML=`<td>${r['사업코드']}</td><td>${r['예산과목']}</td>
            <td><input class="inp-budget" type="text" data-section="cap" data-idx="${i}" value="${fw(a)}" onchange="recalcBudget(this)"></td>
            <td>${fw(b)}</td><td>${fw(c)}</td><td>${fw(d)}</td>
            <td class="${clr(e)}">${fw(e)}</td><td>${br(dr)}</td>
            <td>${fw(f)}</td><td>${fw(g)}</td><td class="${clr(ga)}">${fw(ga)}</td><td>${br(gr)}</td><td>${r['건수']}</td>`;
        tb.appendChild(tr);
    });
    updateTotals('cap');
}

// ═══════════════════════════════════════
// 손익 렌더
// ═══════════════════════════════════════
function renderRevenue(){
    const s=D.revenue.summary;
    document.getElementById('revA').textContent=fm(s['배정예산']);
    document.getElementById('revD').textContent=fm(s['집행실적']);
    document.getElementById('revDR').textContent='집행율 '+fp(s['집행율']);
    document.getElementById('revE').textContent=fm(s['잔액']);
    document.getElementById('revF').textContent=fm(s['진행중공사비']);
    document.getElementById('revG').textContent=fm(s['예상집행']);
    document.getElementById('revGR').textContent='예상집행율 '+fp(s['예상집행율']);
    document.getElementById('revCnt').textContent=s['공사건수']+'건';
    document.getElementById('revOver').textContent='초과 '+s['초과항목']+'건';

    const comp=D.revenue.budget_comparison;
    const lb=[],aV=[],dV=[],gV=[];
    comp.forEach(r=>{if(!r['배정예산_백만'])return;lb.push(sn(r['예산과목'],7));aV.push(r['배정예산_백만']);dV.push(Math.round(r['집행실적']/1e6));gV.push(Math.round(r['예상집행']/1e6))});
    const revMax=Math.max(...aV,...dV,...gV);const revYMax=niceMax(revMax);
    _ch('chRevBar','bar',{labels:lb,datasets:[
        {label:'배정예산(A)',data:aV,backgroundColor:'rgba(59,130,246,.7)',borderColor:'rgba(59,130,246,1)',borderWidth:1,borderRadius:3},
        {label:'집행실적(D)',data:dV,backgroundColor:'rgba(239,68,68,.65)',borderColor:'rgba(239,68,68,1)',borderWidth:1,borderRadius:3},
        {label:'최종예상(G)',data:gV,backgroundColor:'rgba(251,191,36,.6)',borderColor:'rgba(245,158,11,1)',borderWidth:1,borderRadius:3}
    ]},{scales:{y:{beginAtZero:true,max:revYMax,afterBuildTicks(axis){const t=[0,50,100];for(let v=200;v<=axis.max;v+=200)t.push(v);axis.ticks=t.map(v=>({value:v}))},ticks:{font:{size:10}},title:{display:true,text:'백만원'}}}});

    const tb=document.querySelector('#tRevMain tbody');tb.innerHTML='';
    comp.forEach((r,i)=>{
        const a=r['배정예산'],b=r['소비금액'],c=r['약정금액'],d=r['집행실적'],e=r['잔액'],f=r['진행중공사비'],g=r['예상집행'],ga=r['예상잔액'],dr=r['집행율'],gr=r['예상집행율'];
        const tr=document.createElement('tr');
        tr.innerHTML=`<td>${r['사업코드']}</td><td>${r['예산과목']}</td>
            <td><input class="inp-budget" type="text" data-section="rev" data-idx="${i}" value="${fw(a)}" onchange="recalcBudget(this)"></td>
            <td>${fw(b)}</td><td>${fw(c)}</td><td>${fw(d)}</td>
            <td class="${clr(e)}">${fw(e)}</td><td>${br(dr)}</td>
            <td>${fw(f)}</td><td>${fw(g)}</td><td class="${clr(ga)}">${fw(ga)}</td><td>${br(gr)}</td><td>${r['건수']}</td>`;
        tb.appendChild(tr);
    });
    updateTotals('rev');
}

// ═══════════════════════════════════════
// 배정예산 입력 → 자동 재계산 (자본/손익 공통)
// ═══════════════════════════════════════
function recalcBudget(inp){
    const sec=inp.dataset.section; // 'cap' or 'rev'
    const idx=parseInt(inp.dataset.idx);
    const newA=parseFloat(inp.value.replace(/,/g,''))||0;
    const dataKey=sec==='cap'?'capital':'revenue';
    const r=D[dataKey].budget_comparison[idx];
    r['배정예산']=newA;
    r['배정예산_백만']=newA/1e6;
    const d=r['집행실적'],g=r['예상집행'];
    r['잔액']=newA-d;
    r['예상잔액']=newA-g;
    r['집행율']=newA?+(d/newA*100).toFixed(1):0;
    r['예상집행율']=newA?+(g/newA*100).toFixed(1):0;
    r['상태']=r['예상집행율']>100?'초과':r['예상집행율']>70?'양호':'미달';
    inp.value=fw(newA);
    // 행 업데이트
    const tr=inp.closest('tr');const cells=tr.querySelectorAll('td');
    cells[6].className=clr(r['잔액']);cells[6].textContent=fw(r['잔액']);
    cells[7].innerHTML=br(r['집행율']);
    cells[10].className=clr(r['예상잔액']);cells[10].textContent=fw(r['예상잔액']);
    cells[11].innerHTML=br(r['예상집행율']);
    updateTotals(sec);
    updateSummaryCards(sec);
}

function updateTotals(sec){
    const dataKey=sec==='cap'?'capital':'revenue';
    const comp=D[dataKey].budget_comparison;
    const tA=comp.reduce((s,r)=>s+r['배정예산'],0);
    const tB=comp.reduce((s,r)=>s+r['소비금액'],0);
    const tC=comp.reduce((s,r)=>s+r['약정금액'],0);
    const tD=comp.reduce((s,r)=>s+r['집행실적'],0);
    const tE=tA-tD;
    const tF=comp.reduce((s,r)=>s+r['진행중공사비'],0);
    const tG=comp.reduce((s,r)=>s+r['예상집행'],0);
    const tGA=tA-tG;
    const tN=comp.reduce((s,r)=>s+r['건수'],0);
    const tDR=tA?+(tD/tA*100).toFixed(1):0;
    const tGR=tA?+(tG/tA*100).toFixed(1):0;
    const p=sec;
    document.getElementById(p+'TotA').textContent=fw(tA);
    document.getElementById(p+'TotB').textContent=fw(tB);
    document.getElementById(p+'TotC').textContent=fw(tC);
    document.getElementById(p+'TotD').textContent=fw(tD);
    const eEl=document.getElementById(p+'TotE');eEl.className=clr(tE);eEl.textContent=fw(tE);
    document.getElementById(p+'TotDR').innerHTML=br(tDR);
    document.getElementById(p+'TotF').textContent=fw(tF);
    document.getElementById(p+'TotG').textContent=fw(tG);
    const gaEl=document.getElementById(p+'TotGA');gaEl.className=clr(tGA);gaEl.textContent=fw(tGA);
    document.getElementById(p+'TotGR').innerHTML=br(tGR);
    document.getElementById(p+'TotN').textContent=tN;
}

function updateSummaryCards(sec){
    const dataKey=sec==='cap'?'capital':'revenue';
    const comp=D[dataKey].budget_comparison;
    const totA=comp.reduce((s,r)=>s+r['배정예산'],0);
    const totD=comp.reduce((s,r)=>s+r['집행실적'],0);
    const totF=comp.reduce((s,r)=>s+r['진행중공사비'],0);
    const totG=comp.reduce((s,r)=>s+r['예상집행'],0);
    const sm=D[dataKey].summary;
    sm['배정예산']=totA;sm['집행실적']=totD;sm['잔액']=totA-totD;
    sm['집행율']=totA?+(totD/totA*100).toFixed(1):0;
    sm['진행중공사비']=totF;sm['예상집행']=totG;sm['예상잔액']=totA-totG;
    sm['예상집행율']=totA?+(totG/totA*100).toFixed(1):0;
    sm['초과항목']=comp.filter(r=>r['상태']==='초과').length;
    const p=sec;
    document.getElementById(p+'A').textContent=fm(totA);
    document.getElementById(p+'D').textContent=fm(totD);
    document.getElementById(p+'DR').textContent='집행율 '+fp(sm['집행율']);
    document.getElementById(p+'E').textContent=fm(totA-totD);
    document.getElementById(p+'F').textContent=fm(totF);
    document.getElementById(p+'G').textContent=fm(totG);
    document.getElementById(p+'GR').textContent='예상집행율 '+fp(sm['예상집행율']);
    if(document.getElementById(p+'Over'))document.getElementById(p+'Over').textContent='초과 '+sm['초과항목']+'건';
}

// ═══════════════════════════════════════
// 항목 추가
// ═══════════════════════════════════════
function addRow(sec){
    const dataKey=sec==='cap'?'capital':'revenue';
    const comp=D[dataKey].budget_comparison;
    const tId=sec==='cap'?'tCapMain':'tRevMain';
    const idx=comp.length;
    const newItem={'사업코드':'','예산과목':'','배정예산_백만':0,'배정예산':0,'소비금액':0,'약정금액':0,'집행실적':0,'잔액':0,'집행율':0,'진행중공사비':0,'예상집행':0,'예상잔액':0,'예상집행율':0,'건수':0,'상태':'미달'};
    comp.push(newItem);
    const tb=document.querySelector('#'+tId+' tbody');
    const tr=document.createElement('tr');
    tr.innerHTML=`<td><input class="inp-code" type="text" value="" placeholder="사업코드" onchange="this.closest('tr');D['${dataKey}'].budget_comparison[${idx}]['사업코드']=this.value"></td>`
        +`<td><input class="inp-code" type="text" value="" placeholder="사업명" onchange="D['${dataKey}'].budget_comparison[${idx}]['예산과목']=this.value"></td>`
        +`<td><input class="inp-budget" type="text" data-section="${sec}" data-idx="${idx}" value="0" onchange="recalcBudget(this)"></td>`
        +`<td>${fw(0)}</td><td>${fw(0)}</td><td>${fw(0)}</td>`
        +`<td class="pos">${fw(0)}</td><td>${br(0)}</td>`
        +`<td>${fw(0)}</td><td>${fw(0)}</td><td class="pos">${fw(0)}</td><td>${br(0)}</td><td>0</td>`;
    tb.appendChild(tr);
    tr.scrollIntoView({behavior:'smooth',block:'center'});
    tr.querySelector('input').focus();
}

// ═══════════════════════════════════════
// 공사목록
// ═══════════════════════════════════════
function renderProjects(){
    const rows=D.projects;
    setupPT(rows,'tCapProj','capSrch','capFilt','자본예산과목','설계_자본','기성_자본','예정_자본');
    setupPT(rows,'tRevProj','revSrch','revFilt','손익예산과목','설계_손익','기성_손익','예정_손익');
}
function setupPT(rows,tId,srchId,filtId,catKey,dK,pK,eK){
    const tb=document.querySelector('#'+tId+' tbody');
    function render(flt,srch){
        tb.innerHTML='';
        rows.filter(r=>{
            if(flt&&!r['공사상태'].includes(flt))return false;
            if(srch){const s=srch.toLowerCase();return r['공사번호'].toLowerCase().includes(s)||r['공사업체'].toLowerCase().includes(s)||r[catKey].toLowerCase().includes(s)}
            return true;
        }).forEach((r,idx)=>{
            const rate=r[dK]?(r[pK]/r[dK]*100):0;
            const tr=document.createElement('tr');
            tr.innerHTML=`<td>${idx+1}</td><td>${r['공사번호']}</td><td>${r[catKey]}</td><td>${r['공사업체']}</td><td>${stBg(r['공사상태'])}</td><td>${r['착공일']}</td><td>${fw(r[dK])}</td><td>${fw(r[pK])}</td><td>${fw(r[eK])}</td><td>${br(rate)}</td>`;
            tb.appendChild(tr);
        });
    }
    render('','');
    document.getElementById(srchId).addEventListener('input',e=>render(document.getElementById(filtId).value,e.target.value));
    document.getElementById(filtId).addEventListener('change',e=>render(e.target.value,document.getElementById(srchId).value));
}

window.addEventListener('DOMContentLoaded',async()=>{
    const now=new Date();const y=now.getFullYear();const m=String(now.getMonth()+1).padStart(2,'0');const d=String(now.getDate()).padStart(2,'0');const wd=['일','월','화','수','목','금','토'][now.getDay()];
    document.getElementById('today').textContent=y+'.'+m+'.'+d+' ('+wd+')';
    try{const r=await fetch('/api/default');const j=await r.json();if(!j.error){D=j;document.getElementById('fn').textContent='기본 데이터 로드됨';document.getElementById('refreshBtn').style.display='inline-block';renderAll()}}catch(e){}
});
</script>
</body>
</html>
"""

# ──────────────────────────────────────────────
# 5. Flask 라우트
# ──────────────────────────────────────────────
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({'error': '파일명이 비어있습니다.'}), 400
    filepath = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(filepath)
    try:
        return jsonify(parse_and_analyze(filepath))
    except Exception as e:
        return jsonify({'error': f'분석 오류: {str(e)}'}), 500


@app.route('/api/default')
def api_default():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '예산관리 엑셀.xlsx')
    if not os.path.exists(path):
        return jsonify({'error': '기본 파일 없음'}), 404
    try:
        return jsonify(parse_and_analyze(path))
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5005)
