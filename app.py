# -*- coding: utf-8 -*-
import os, json
from flask import Flask, request, jsonify, render_template_string
import pandas as pd
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
LAST_FILE = None
BUDGET_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'budgets.json')

# ──────────────────────────────────────────────
# 1. 배정예산 데이터 (백만원)
# ──────────────────────────────────────────────
BUDGET_CAPITAL = {
    #  사업명: (자금운용사업코드, 배정예산(백만원))
    '가공저압설비보강': ('310400201180203', 0),
    '취약선로보강': ('310400203200089', 0),
    '가공고압설비보강': ('310400208200094', 0),
    '노후 가공변압기 교체': ('310400212200088', 0),
    '과부하 가공변압기 교체': ('310400214200084', 0),
    '강도부족전주 교체': ('310400225200084', 0),
    '가공개폐기 교체': ('310400226200081', 0),
    'PCBS근절 변압기 교체': ('70H120224128', 0),
    '차량충돌복구공사': ('310400218200086', 0),
    '고장변압기 교체': ('310400213200087', 0),
    '재해설비 피해복구': ('70H120234279', 0),
    '가공설비고장복구': ('310400219200086', 0),
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
    '배전경상': ('S10030', 0),
    '내선경상': ('S10040', 0),
    '가공배전설비진단': ('S20110', 0),
    '배전수목전지': ('S20140', 0),
    '배전조류고장예방': ('S20150', 0),
    '배전염진해낙뢰': ('S20160', 0),
    '가공배전기타계획(주요설비)': ('S20176', 0),
    '가공배전기타계획(기타설비)': ('S20180', 0),
    '계기함정비': ('S20190', 0),
    '자본연계(공급_선진)': ('S20240', 0),
    '자본연계(보강/가공)': ('S20245', 0),
    '지중배전기타점검(주요설비)': ('S20265', 0),
    '배전기자재수리': ('S20290', 0),
    '자본연계(신규)': ('S20220', 0),
    '자본연계(지장)': ('S20230', 0),
    '자본연계(보강/내선)': ('S20247', 0),
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
# 3-A. AI 분석 함수 (통계/규칙 기반)
# ──────────────────────────────────────────────
def _predict_yearend(comparison, budget_dict):
    """경과 월수 기반 연말 예상 집행액 추정"""
    now = datetime.now()
    month = now.month
    elapsed = month / 12  # 연간 경과 비율

    if elapsed <= 0:
        elapsed = 1 / 12

    confidence = '높음' if month >= 9 else ('보통' if month >= 6 else '낮음')

    predictions = []
    total_budget = sum(v[1] * 1e6 for v in budget_dict.values())
    total_current_exec = 0
    total_predicted = 0

    for r in comparison:
        a = r['배정예산']
        d = r['집행실적']
        f = r['진행중공사비']
        total_current_exec += d

        # 연말 예측: 현재 집행실적을 경과비율로 나눈 값
        if d > 0:
            projected = round(d / elapsed)
        else:
            projected = f  # 집행실적 없으면 진행중공사비만

        # 진행중공사비 반영 (최소 집행실적 + 진행중공사비)
        projected = max(projected, d + f)
        total_predicted += projected

        risk = '초과위험' if a > 0 and projected > a else ('주의' if a > 0 and projected > a * 0.9 else '정상')

        predictions.append({
            '예산과목': r['예산과목'],
            '배정예산': a,
            '현재집행': d,
            '연말예측': projected,
            '예측집행율': round(projected / a * 100, 1) if a else 0,
            '위험도': risk,
        })

    return {
        'month': month,
        'elapsed_pct': round(elapsed * 100, 1),
        'confidence': confidence,
        'total_budget': total_budget,
        'total_current_exec': total_current_exec,
        'total_predicted': total_predicted,
        'predicted_rate': round(total_predicted / total_budget * 100, 1) if total_budget else 0,
        'items': predictions,
    }


def _detect_anomalies(comparison, budget_dict):
    """예산 집행 패턴 이상 탐지"""
    now = datetime.now()
    month = now.month
    expected_rate = month / 12 * 100  # 시간 비례 기대 집행율

    anomalies = []

    for r in comparison:
        a = r['배정예산']
        name = r['예산과목']
        exec_rate = r['집행율']
        fc_rate = r['예상집행율']

        if a <= 0:
            continue

        # 시간 대비 과다 집행
        if exec_rate > expected_rate + 15:
            anomalies.append({
                'level': 'danger',
                'category': name,
                'message': f'시간 대비 과다 집행 (집행율 {exec_rate}% vs 기대 {expected_rate:.0f}%)',
                'detail': f'현재 {month}월 기준 기대집행율 {expected_rate:.0f}% 대비 {exec_rate - expected_rate:.1f}%p 초과',
            })

        # 예산 초과 위험
        if fc_rate > 100:
            anomalies.append({
                'level': 'danger',
                'category': name,
                'message': f'예산 초과 (예상집행율 {fc_rate}%)',
                'detail': f'예상집행이 배정예산을 {fc_rate - 100:.1f}% 초과',
            })
        elif fc_rate > 90:
            anomalies.append({
                'level': 'warning',
                'category': name,
                'message': f'예산 초과 주의 (예상집행율 {fc_rate}%)',
                'detail': f'배정예산 소진율 {fc_rate}%로 초과 가능성 있음',
            })

        # 시간 대비 과소 집행 (하반기부터 체크)
        if month >= 6 and exec_rate < expected_rate - 30 and a >= 1_000_000:
            anomalies.append({
                'level': 'info',
                'category': name,
                'message': f'집행 부진 (집행율 {exec_rate}% vs 기대 {expected_rate:.0f}%)',
                'detail': f'배정예산 대비 집행이 느림 - 하반기 집중 집행 필요 가능성',
            })

    # 위험도순 정렬
    order = {'danger': 0, 'warning': 1, 'info': 2}
    anomalies.sort(key=lambda x: order.get(x['level'], 9))
    return anomalies


def _generate_report(cap_comp, rev_comp, cap_pred, rev_pred, cap_anomalies, rev_anomalies,
                     budget_cap, budget_rev):
    """예산 현황 보고서 텍스트 생성"""
    now = datetime.now()
    date_str = now.strftime('%Y년 %m월 %d일')
    month = now.month
    elapsed_pct = round(month / 12 * 100, 1)

    lines = []
    lines.append('=' * 50)
    lines.append('    배전공사 예산 집행 현황 보고서')
    lines.append('=' * 50)
    lines.append(f'기준일: {date_str} (연간 진행율 {elapsed_pct}%)')
    lines.append('')

    # 자본
    tc_budget = sum(v[1] * 1e6 for v in budget_cap.values())
    tc_exec = sum(r['집행실적'] for r in cap_comp)
    tc_forecast = sum(r['예상집행'] for r in cap_comp)
    tc_rate = round(tc_exec / tc_budget * 100, 1) if tc_budget else 0
    tc_fc_rate = round(tc_forecast / tc_budget * 100, 1) if tc_budget else 0

    lines.append('■ 자본예산 현황')
    lines.append(f'  - 배정예산: {tc_budget / 1e8:.1f}억원')
    lines.append(f'  - 집행실적: {tc_exec / 1e8:.1f}억원 (집행율 {tc_rate}%)')
    lines.append(f'  - 예상집행: {tc_forecast / 1e8:.1f}억원 (예상집행율 {tc_fc_rate}%)')
    if cap_pred:
        lines.append(f'  - 연말 예측: {cap_pred["total_predicted"] / 1e8:.1f}억원 (신뢰도: {cap_pred["confidence"]})')

    # 자본 주요 항목
    cap_over = [r for r in cap_comp if r['예상집행율'] > 90 and r['배정예산'] > 0]
    if cap_over:
        lines.append(f'  - 주의 항목 ({len(cap_over)}건):')
        for r in cap_over:
            lines.append(f'    * {r["예산과목"]}: 예상집행율 {r["예상집행율"]}%')
    lines.append('')

    # 손익
    tr_budget = sum(v[1] * 1e6 for v in budget_rev.values())
    tr_exec = sum(r['집행실적'] for r in rev_comp)
    tr_forecast = sum(r['예상집행'] for r in rev_comp)
    tr_rate = round(tr_exec / tr_budget * 100, 1) if tr_budget else 0
    tr_fc_rate = round(tr_forecast / tr_budget * 100, 1) if tr_budget else 0

    lines.append('■ 손익예산 현황')
    lines.append(f'  - 배정예산: {tr_budget / 1e8:.1f}억원')
    lines.append(f'  - 집행실적: {tr_exec / 1e8:.1f}억원 (집행율 {tr_rate}%)')
    lines.append(f'  - 예상집행: {tr_forecast / 1e8:.1f}억원 (예상집행율 {tr_fc_rate}%)')
    if rev_pred:
        lines.append(f'  - 연말 예측: {rev_pred["total_predicted"] / 1e8:.1f}억원 (신뢰도: {rev_pred["confidence"]})')

    rev_over = [r for r in rev_comp if r['예상집행율'] > 90 and r['배정예산'] > 0]
    if rev_over:
        lines.append(f'  - 주의 항목 ({len(rev_over)}건):')
        for r in rev_over:
            lines.append(f'    * {r["예산과목"]}: 예상집행율 {r["예상집행율"]}%')
    lines.append('')

    # 이상 탐지
    all_anomalies = cap_anomalies + rev_anomalies
    if all_anomalies:
        lines.append(f'■ 이상 탐지 결과 ({len(all_anomalies)}건)')
        level_label = {'danger': '위험', 'warning': '주의', 'info': '정보'}
        for a in all_anomalies:
            lbl = level_label.get(a['level'], '정보')
            section = '자본' if a in cap_anomalies else '손익'
            lines.append(f'  [{lbl}] [{section}] {a["category"]}: {a["message"]}')
    else:
        lines.append('■ 이상 탐지 결과: 특이사항 없음')
    lines.append('')

    # 종합 의견
    lines.append('■ 종합 의견')
    danger_cnt = sum(1 for a in all_anomalies if a['level'] == 'danger')
    warning_cnt = sum(1 for a in all_anomalies if a['level'] == 'warning')
    if danger_cnt > 0:
        lines.append(f'  - 위험 항목 {danger_cnt}건 감지. 예산 초과 방지를 위한 긴급 점검 필요.')
    if warning_cnt > 0:
        lines.append(f'  - 주의 항목 {warning_cnt}건. 하반기 집행 계획 재검토 권고.')
    if danger_cnt == 0 and warning_cnt == 0:
        lines.append('  - 전체적으로 예산 집행이 정상 범위 내에 있습니다.')

    remaining_months = 12 - month
    if remaining_months > 0:
        lines.append(f'  - 잔여 기간: {remaining_months}개월')

    lines.append('')
    lines.append('=' * 50)
    lines.append(f'※ 본 보고서는 {date_str} 기준 자동 생성되었습니다.')

    return '\n'.join(lines)


# ──────────────────────────────────────────────
# 3-A2. 추가 AI 분석 함수
# ──────────────────────────────────────────────
def _burndown_forecast(comparison, budget_dict):
    """월별 예산 소진 예측 곡선 데이터 생성"""
    now = datetime.now()
    month = now.month
    total_budget = sum(r['배정예산'] for r in comparison)
    total_exec = sum(r['집행실적'] for r in comparison)
    total_progress = sum(r['진행중공사비'] for r in comparison)

    actual_monthly = []
    for m in range(1, 13):
        if m <= month:
            actual_monthly.append(round(total_exec * m / month) if month > 0 else 0)
        else:
            actual_monthly.append(None)

    monthly_rate = total_exec / month if month > 0 else 0
    projected_monthly = [round(monthly_rate * m) for m in range(1, 13)]

    committed_total = total_exec + total_progress
    committed_rate = committed_total / month if month > 0 else 0
    committed_monthly = [round(committed_rate * m) for m in range(1, 13)]

    budget_line = [round(total_budget)] * 12
    even_monthly = [round(total_budget * m / 12) for m in range(1, 13)]

    exhaustion_month = None
    if monthly_rate > 0:
        em = total_budget / monthly_rate
        if em <= 12:
            exhaustion_month = round(em, 1)

    return {
        'months': list(range(1, 13)),
        'current_month': month,
        'total_budget': total_budget,
        'total_exec': total_exec,
        'actual_monthly': actual_monthly,
        'projected_monthly': projected_monthly,
        'committed_monthly': committed_monthly,
        'budget_line': budget_line,
        'even_monthly': even_monthly,
        'exhaustion_month': exhaustion_month,
        'monthly_rate': monthly_rate,
    }


def _delay_risk_scores(projects):
    """공사별 지연 리스크 점수 산출 (0-100)"""
    now = datetime.now()
    scored = []

    for p in projects:
        score = 0
        factors = []
        status = p.get('공사상태', '')

        if '완료' in status:
            scored.append({
                '공사번호': p['공사번호'], '공사업체': p.get('공사업체', ''),
                '공사상태': status, '착공일': p.get('착공일', ''), '준공일': p.get('준공일', ''),
                'risk_score': 0, 'risk_level': 'low', 'factors': ['완료'],
                '자본예산과목': p.get('자본예산과목', ''), '손익예산과목': p.get('손익예산과목', ''),
                '설계_자본': p.get('설계_자본', 0), '기성_자본': p.get('기성_자본', 0),
                '설계_손익': p.get('설계_손익', 0), '기성_손익': p.get('기성_손익', 0),
            })
            continue

        if '중지' in status:
            score += 30
            factors.append('공사중지(+30)')
        else:
            start_str = p.get('착공일', '')
            end_str = p.get('준공일', '') or p.get('현장시공완료일', '')
            if start_str and end_str:
                try:
                    start_dt = datetime.strptime(start_str[:10], '%Y-%m-%d')
                    end_dt = datetime.strptime(end_str[:10], '%Y-%m-%d')
                    total_days = (end_dt - start_dt).days
                    elapsed_days = (now - start_dt).days
                    if total_days > 0:
                        expected_progress = min(1.0, elapsed_days / total_days)
                        total_design = p.get('설계_자본', 0) + p.get('설계_손익', 0)
                        total_paid = p.get('기성_자본', 0) + p.get('기성_손익', 0)
                        actual_progress = total_paid / total_design if total_design > 0 else 0
                        gap = expected_progress - actual_progress
                        if gap > 0.3:
                            score += 40; factors.append(f'진행율격차({gap:.0%})(+40)')
                        elif gap > 0.15:
                            score += 25; factors.append(f'진행율격차({gap:.0%})(+25)')
                        elif gap > 0:
                            score += 10; factors.append(f'약간지연({gap:.0%})(+10)')
                        if elapsed_days > total_days:
                            overdue_pts = min(30, round((elapsed_days - total_days) / total_days * 60))
                            score += overdue_pts; factors.append(f'기한초과(+{overdue_pts})')
                except (ValueError, TypeError):
                    pass
            elif start_str and not end_str:
                score += 10; factors.append('준공일미설정(+10)')

        score = min(100, max(0, score))
        risk_level = 'high' if score >= 60 else ('medium' if score >= 30 else 'low')
        scored.append({
            '공사번호': p['공사번호'], '공사업체': p.get('공사업체', ''),
            '공사상태': status, '착공일': p.get('착공일', ''), '준공일': p.get('준공일', ''),
            'risk_score': score, 'risk_level': risk_level, 'factors': factors,
            '자본예산과목': p.get('자본예산과목', ''), '손익예산과목': p.get('손익예산과목', ''),
            '설계_자본': p.get('설계_자본', 0), '기성_자본': p.get('기성_자본', 0),
            '설계_손익': p.get('설계_손익', 0), '기성_손익': p.get('기성_손익', 0),
        })

    scored.sort(key=lambda x: x['risk_score'], reverse=True)
    high_count = sum(1 for s in scored if s['risk_level'] == 'high')
    med_count = sum(1 for s in scored if s['risk_level'] == 'medium')
    avg_score = round(sum(s['risk_score'] for s in scored) / len(scored), 1) if scored else 0
    return {
        'items': scored[:30],
        'total_projects': len(scored),
        'high_risk': high_count, 'medium_risk': med_count, 'avg_score': avg_score,
    }


def _whatif_baseline(comparison, budget_dict):
    """What-if 시뮬레이션 기준 데이터"""
    now = datetime.now()
    month = now.month
    remaining = 12 - month
    total_budget = sum(r['배정예산'] for r in comparison)
    total_exec = sum(r['집행실적'] for r in comparison)
    total_forecast = sum(r['예상집행'] for r in comparison)
    monthly_rate = total_exec / month if month > 0 else 0
    return {
        'current_month': month, 'remaining_months': remaining,
        'total_budget': total_budget, 'total_exec': total_exec,
        'total_forecast': total_forecast, 'monthly_rate': monthly_rate,
        'current_yearend_projected': round(monthly_rate * 12),
        'current_yearend_rate': round(monthly_rate * 12 / total_budget * 100, 1) if total_budget else 0,
        'target_60': total_budget * 0.6, 'target_90': total_budget * 0.9,
        'exec_rate': round(total_exec / total_budget * 100, 1) if total_budget else 0,
    }


def _reallocation_recommendations(comparison):
    """예산 재배분 추천: 잉여 항목 → 부족 항목 이전 제안 (단일 섹션)"""
    recommendations = []
    surplus_items, deficit_items = [], []
    for r in comparison:
        budget = r['배정예산']
        if budget <= 0:
            continue
        fc_rate = r['예상집행율']
        remaining = budget - r['예상집행']
        if fc_rate < 70 and remaining > 1_000_000:
            surplus_items.append({
                '예산과목': r['예산과목'], '배정예산': budget,
                '예상집행': r['예상집행'], '예상집행율': fc_rate, '잉여액': remaining,
            })
        if fc_rate > 90:
            shortfall = r['예상집행'] - budget if r['예상집행'] > budget else 0
            deficit_items.append({
                '예산과목': r['예산과목'], '배정예산': budget,
                '예상집행': r['예상집행'], '예상집행율': fc_rate, '부족액': shortfall,
            })
    surplus_items.sort(key=lambda x: x['잉여액'], reverse=True)
    deficit_items.sort(key=lambda x: x['예상집행율'], reverse=True)
    for deficit in deficit_items:
        if deficit['부족액'] <= 0:
            continue
        remaining_need = deficit['부족액']
        for surplus in surplus_items:
            if surplus['잉여액'] <= 0:
                continue
            transfer = min(remaining_need, surplus['잉여액'] * 0.5)
            if transfer < 500_000:
                continue
            new_surplus_rate = round(surplus['예상집행'] / (surplus['배정예산'] - transfer) * 100, 1) if (surplus['배정예산'] - transfer) > 0 else 0
            new_deficit_rate = round(deficit['예상집행'] / (deficit['배정예산'] + transfer) * 100, 1) if (deficit['배정예산'] + transfer) > 0 else 0
            recommendations.append({
                'source': surplus['예산과목'],
                'source_budget': surplus['배정예산'], 'source_fc_rate': surplus['예상집행율'],
                'source_new_rate': new_surplus_rate,
                'target': deficit['예산과목'],
                'target_budget': deficit['배정예산'], 'target_fc_rate': deficit['예상집행율'],
                'target_new_rate': new_deficit_rate,
                'amount': round(transfer),
                'reason': f"{surplus['예산과목']} 잔액 {surplus['잉여액']/1e8:.1f}억 → {deficit['예산과목']} 부족분 {deficit['부족액']/1e8:.1f}억 이전",
            })
            remaining_need -= transfer
            surplus['잉여액'] -= transfer
            if remaining_need <= 0:
                break
    total_surplus = sum(r['amount'] for r in recommendations)
    return {'recommendations': recommendations[:10], 'total_transferable': total_surplus, 'count': len(recommendations)}


# ──────────────────────────────────────────────
# 3-B. 파싱 + 분석
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

    # ── 헤더 행 찾기 ──
    header_row = 0
    for i in range(min(5, len(df))):
        if pd.notna(df.iloc[i, 0]) and '순번' in str(df.iloc[i, 0]):
            header_row = i
            break

    # ── 서브헤더 확인 (더블 헤더) ──
    data_start = header_row + 1
    if data_start < len(df) and pd.notna(df.iloc[data_start, 0]) and '순번' in str(df.iloc[data_start, 0]):
        data_start += 1

    # ── 컬럼 자동 감지: 헤더에서 키워드로 컬럼 인덱스 매핑 ──
    def _find_col(keyword, row_range=None):
        if row_range is None:
            row_range = range(header_row, data_start)
        for r in row_range:
            for c in range(df.shape[1]):
                v = str(df.iloc[r, c]).strip() if pd.notna(df.iloc[r, c]) else ''
                if keyword in v:
                    return c
        return None

    col_cno = _find_col('공사번호') or 1
    col_char = _find_col('공사성격') or 4
    col_company = _find_col('협력회사') or _find_col('시공회사') or _find_col('공사업체') or 6
    col_cat_cap = _find_col('자본예산과목') or 9
    col_cat_rev = _find_col('손익예산과목') or 10
    col_status = _find_col('공사 상태') or _find_col('공사상태') or 22
    col_start_date = _find_col('착공') or 14
    col_site_done = _find_col('현장시공') or _find_col('시공완료') or 15
    col_completion = _find_col('준공검사') or _find_col('준공계') or 16

    # 금액 컬럼: 총공사비(자본/수익) 감지
    # 서브헤더에서 '자본'/'수익' 찾기 (총공사비 그룹 하위)
    col_cost_cap = None
    col_cost_rev = None
    sub_row = data_start - 1
    for c in range(df.shape[1]):
        h0 = str(df.iloc[header_row, c]).strip() if pd.notna(df.iloc[header_row, c]) else ''
        h1 = str(df.iloc[sub_row, c]).strip() if sub_row > header_row and pd.notna(df.iloc[sub_row, c]) else ''
        if '총공사비' in h0 and '자본' in h1:
            col_cost_cap = c
        elif '총공사비' in h0 and ('수익' in h1 or '손익' in h1):
            col_cost_rev = c

    # 총공사비 자본/수익 못 찾으면 구형 포맷 (41, 47 등) 시도
    is_old_format = col_cost_cap is None
    if is_old_format:
        col_cat_cap = 13
        col_cat_rev = 14
        col_status = 26
        col_start_date = 18

    projects = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        cno = row[col_cno] if col_cno < df.shape[1] else None
        if pd.isna(cno) or str(cno).strip() == '':
            continue

        cat_capital = str(row[col_cat_cap]).strip() if col_cat_cap < df.shape[1] and pd.notna(row.get(col_cat_cap)) else '미분류'
        # 손익예산과목: "배전경상-경남/통영/전력공급팀" → "배전경상"
        raw_rev = str(row[col_cat_rev]).strip() if col_cat_rev < df.shape[1] and pd.notna(row.get(col_cat_rev)) else '미분류'
        cat_revenue = raw_rev.split('-')[0].strip() if '-' in raw_rev else raw_rev
        status = str(row[col_status]).strip() if col_status < df.shape[1] and pd.notna(row.get(col_status)) else '미확인'

        if is_old_format:
            # 구형 포맷 (예산관리 엑셀.xlsx)
            cost_cap = _num(row[41]) if 41 < df.shape[1] else 0
            paid_cap = _num(row[47]) if 47 < df.shape[1] else 0
            est_cap = _num(row[51]) if 51 < df.shape[1] else 0
            cost_rev = _num(row[42]) if 42 < df.shape[1] else 0
            paid_rev = _num(row[48]) if 48 < df.shape[1] else 0
            est_rev = _num(row[52]) if 52 < df.shape[1] else 0
        else:
            # 신규 포맷 (공사관리대장조회.xlsx)
            cost_cap = _num(row[col_cost_cap]) if col_cost_cap else 0
            cost_rev = _num(row[col_cost_rev]) if col_cost_rev else 0
            if '완료' in status:
                paid_cap, paid_rev = cost_cap, cost_rev
                est_cap, est_rev = 0, 0
            else:
                paid_cap, paid_rev = 0, 0
                est_cap, est_rev = cost_cap, cost_rev

        projects.append({
            '공사번호': str(cno).strip(),
            '공사성격': str(row[col_char]).strip() if col_char < df.shape[1] and pd.notna(row.get(col_char)) else '',
            '공사업체': str(row[col_company]).strip() if col_company < df.shape[1] and pd.notna(row.get(col_company)) else '',
            '자본예산과목': cat_capital,
            '손익예산과목': cat_revenue,
            '총공사비_설계': _num(row[11]) if 11 < df.shape[1] else 0,
            '도급비_계약': _num(row[13]) if 13 < df.shape[1] else 0,
            '착공일': _date_str(row[col_start_date]) if col_start_date < df.shape[1] else '',
            '현장시공완료일': _date_str(row[col_site_done]) if col_site_done < df.shape[1] else '',
            '준공일': _date_str(row[col_completion]) if col_completion < df.shape[1] else '',
            '공사상태': status,
            '설계_자본': cost_cap,
            '기성_자본': paid_cap,
            '예정_자본': est_cap,
            '설계_손익': cost_rev,
            '기성_손익': paid_rev,
            '예정_손익': est_rev,
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

    # ── AI 분석 ──
    cap_pred = _predict_yearend(cap_comparison, BUDGET_CAPITAL)
    rev_pred = _predict_yearend(rev_comparison, BUDGET_REVENUE)
    cap_anomalies = _detect_anomalies(cap_comparison, BUDGET_CAPITAL)
    rev_anomalies = _detect_anomalies(rev_comparison, BUDGET_REVENUE)
    report = _generate_report(cap_comparison, rev_comparison,
                               cap_pred, rev_pred,
                               cap_anomalies, rev_anomalies,
                               BUDGET_CAPITAL, BUDGET_REVENUE)
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
        'ai_analysis': {
            'capital': {
                'predictions': cap_pred, 'anomalies': cap_anomalies,
            },
            'revenue': {
                'predictions': rev_pred, 'anomalies': rev_anomalies,
            },
            'report': report,
        },
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

/* ── Project table ── */
#tCapProj td:nth-child(n+9),#tCapProj th:nth-child(n+9),#tRevProj td:nth-child(n+9),#tRevProj th:nth-child(n+9){font-size:12px}

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
.dbl-code{cursor:default;user-select:none}
.dbl-code:hover{background:#EFF6FF;cursor:pointer}
.dbl-code:hover::after{content:' ✎';font-size:9px;color:var(--navy2);opacity:.5}

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

/* ── Empty state ── */
.empty-state{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:60vh;color:var(--text3);gap:16px}
.empty-state .es-icon{font-size:56px;opacity:.4}
.empty-state .es-title{font-size:20px;font-weight:700;color:var(--text2)}
.empty-state .es-desc{font-size:13px;color:var(--text3);text-align:center;line-height:1.6}
.empty-state .es-btn{margin-top:8px;padding:10px 28px;background:var(--orange);color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;transition:background .15s}
.empty-state .es-btn:hover{background:var(--orange-light)}

/* ── AI Panel ── */
.ai-panel{background:linear-gradient(135deg,#F0F4FF 0%,#EBF5FF 50%,#F5F0FF 100%);border:1px solid #C7D2FE;border-radius:10px;padding:18px 20px;margin:16px 0;position:relative;overflow:hidden}
.ai-panel::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:linear-gradient(180deg,#6366F1,#8B5CF6)}
.ai-panel .ai-header{display:flex;align-items:center;gap:8px;margin-bottom:12px;font-size:13px;font-weight:700;color:#4338CA}
.ai-panel .ai-header .ai-icon{width:22px;height:22px;background:linear-gradient(135deg,#6366F1,#8B5CF6);border-radius:6px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:800}
.ai-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:14px}
.ai-stat{background:#fff;border-radius:8px;padding:10px 14px;border:1px solid #E0E7FF}
.ai-stat .ai-label{font-size:10px;color:var(--text3);font-weight:600;margin-bottom:2px}
.ai-stat .ai-value{font-size:16px;font-weight:800;color:var(--text)}
.ai-stat .ai-sub{font-size:10px;color:var(--text3);margin-top:2px}
.ai-alerts{display:flex;flex-direction:column;gap:6px}
.ai-alert{display:flex;align-items:flex-start;gap:8px;padding:8px 12px;border-radius:6px;font-size:11px;line-height:1.5}
.ai-alert.danger{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.ai-alert.warning{background:#FFFBEB;border:1px solid #FDE68A;color:#92400E}
.ai-alert.info{background:#EFF6FF;border:1px solid #BFDBFE;color:#1E40AF}
.ai-alert .ai-alert-icon{flex-shrink:0;width:18px;height:18px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;color:#fff}
.ai-alert.danger .ai-alert-icon{background:#EF4444}
.ai-alert.warning .ai-alert-icon{background:#F59E0B}
.ai-alert.info .ai-alert-icon{background:#3B82F6}
.ai-alert .ai-alert-text{flex:1}
.ai-alert .ai-alert-title{font-weight:700;font-size:11px}
.ai-alert .ai-alert-detail{font-size:10px;opacity:.8;margin-top:1px}

/* ── Burndown Chart ── */
.burn-warn{color:var(--red);font-weight:600}


/* ── What-if Simulation ── */
.whatif-panel{background:linear-gradient(135deg,#ECFDF5 0%,#D1FAE5 50%,#F0FDF4 100%) !important;border-color:#A7F3D0 !important}
.whatif-panel::before{background:linear-gradient(180deg,#10B981,#059669) !important}
.whatif-controls{display:flex;align-items:center;gap:12px;margin-bottom:14px;padding:10px 14px;background:#fff;border-radius:8px;border:1px solid #A7F3D0}
.whatif-controls label{font-size:11px;font-weight:700;color:#065F46;white-space:nowrap}
.whatif-controls input[type=range]{flex:1;height:6px;-webkit-appearance:none;appearance:none;background:#D1FAE5;border-radius:3px;outline:none}
.whatif-controls input[type=range]::-webkit-slider-thumb{-webkit-appearance:none;width:18px;height:18px;background:#10B981;border-radius:50%;cursor:pointer;border:2px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,.2)}
.whatif-pct-label{font-size:16px;font-weight:800;color:#065F46;min-width:45px;text-align:right}
.whatif-chart-wrap{margin:12px 0}
.whatif-message{padding:8px 12px;border-radius:6px;font-size:11px;margin-top:8px}
.whatif-msg-ok{background:#D1FAE5;color:#065F46}
.whatif-msg-warn{background:#FEF3C7;color:#92400E}
.whatif-msg-danger{background:#FEE2E2;color:#991B1B}

/* ── Reallocation ── */
.realloc-panel{background:linear-gradient(135deg,#F5F3FF 0%,#EDE9FE 50%,#FDF4FF 100%) !important;border-color:#C4B5FD !important}
.realloc-panel::before{background:linear-gradient(180deg,#7C3AED,#6D28D9) !important}
.realloc-details{background:#fff;border-radius:8px;border:1px solid #DDD6FE;margin-bottom:8px;overflow:hidden}
.realloc-details summary{display:flex;align-items:center;gap:10px;padding:11px 16px;cursor:pointer;list-style:none;user-select:none}
.realloc-details summary::-webkit-details-marker{display:none}
.realloc-toggle{width:18px;height:18px;background:#EDE9FE;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:9px;color:#7C3AED;flex-shrink:0;transition:transform .2s}
.realloc-details[open] .realloc-toggle{transform:rotate(90deg)}
.realloc-details-body{padding:10px 16px 14px;border-top:1px solid #EDE9FE;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.rc-arrow{font-size:18px;color:#7C3AED;font-weight:800;flex-shrink:0}
.rc-box{flex:1;min-width:140px}
.rc-name{font-size:11px;font-weight:700;color:var(--text)}
.rc-detail{font-size:10px;color:var(--text2);margin-top:2px}
.rc-rate-change{font-size:10px;font-weight:700}
.rc-amount{background:#EDE9FE;color:#5B21B6;padding:5px 12px;border-radius:6px;font-size:12px;font-weight:800;white-space:nowrap;flex-shrink:0}
.rc-reason{font-size:10px;color:var(--text3);width:100%;border-top:1px solid #EDE9FE;padding-top:6px;margin-top:4px}

/* ── Report Modal ── */
.modal-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:center}
.modal-overlay.on{display:flex}
.modal-box{background:#fff;border-radius:12px;width:90%;max-width:700px;max-height:85vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.2)}
.modal-head{display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid var(--border)}
.modal-head h2{font-size:15px;color:var(--text);font-weight:700}
.modal-head .modal-actions{display:flex;gap:6px}
.modal-head button{padding:6px 14px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid var(--border);background:#fff;color:var(--text);transition:all .15s}
.modal-head button:hover{background:#F8FAFC}
.modal-head .btn-copy{background:var(--navy2);color:#fff;border-color:var(--navy2)}
.modal-head .btn-copy:hover{background:var(--navy-light)}
.modal-head .btn-close{background:transparent}
.modal-body{padding:20px;overflow-y:auto;flex:1}
.modal-body pre{white-space:pre-wrap;word-break:break-all;font-family:'Malgun Gothic','Apple SD Gothic Neo',monospace;font-size:12px;line-height:1.8;color:var(--text)}

/* ── Early Execution Target ── */
.early-exec{background:linear-gradient(135deg,#FFF7ED 0%,#FEF3C7 100%);border:1px solid #FDE68A;border-radius:10px;padding:16px 20px;margin:16px 0;position:relative;overflow:hidden}
.early-exec::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:linear-gradient(180deg,var(--orange),var(--amber))}
.early-exec .ee-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.early-exec .ee-title{font-size:13px;font-weight:700;color:#92400E;display:flex;align-items:center;gap:6px}
.early-exec .ee-controls{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.early-exec .ee-controls label{font-size:11px;font-weight:600;color:#92400E}
.early-exec .ee-controls input,.early-exec .ee-controls select{padding:4px 8px;border:1px solid #FDE68A;border-radius:5px;font-size:12px;font-weight:700;width:60px;text-align:center;background:#fff}
.early-exec .ee-controls input:focus,.early-exec .ee-controls select:focus{outline:none;border-color:var(--orange);box-shadow:0 0 0 2px rgba(232,115,27,.15)}
.ee-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px}
.ee-item{background:#fff;border-radius:8px;padding:10px 12px;border:1px solid #FDE68A;display:flex;flex-direction:column;gap:4px}
.ee-item .ee-name{font-size:10px;font-weight:700;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ee-item .ee-bar{height:8px;background:#F3F4F6;border-radius:4px;overflow:hidden;position:relative}
.ee-item .ee-bar .ee-fill{height:100%;border-radius:4px;transition:width .3s}
.ee-item .ee-bar .ee-target-line{position:absolute;top:-2px;height:12px;width:2px;background:#991B1B}
.ee-item .ee-nums{display:flex;justify-content:space-between;font-size:10px;color:var(--text3)}
.ee-item .ee-nums .ee-achieved{font-weight:700}
.ee-item.ee-over .ee-fill{background:var(--green)}.ee-item.ee-close .ee-fill{background:var(--amber)}.ee-item.ee-behind .ee-fill{background:var(--red)}
.ee-total{background:#fff;border-radius:8px;padding:12px 16px;border:2px solid var(--orange);margin-top:8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.ee-total .ee-total-label{font-size:12px;font-weight:700;color:#92400E}
.ee-total .ee-total-value{font-size:18px;font-weight:800;color:var(--text)}
.ee-total .ee-total-sub{font-size:10px;color:var(--text3)}

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
<button class="ubtn" id="reportBtn" onclick="openReport()" style="display:none">&#128196; 보고서</button>
<button class="ubtn" id="refreshBtn" onclick="doRefresh()" style="display:none">&#8635; 새로고침</button>
<label class="ubtn ubtn-primary">파일 업로드<input type="file" id="fi" accept=".xlsx,.xls" style="display:none"></label>
</div>
</header>
<div class="wrap">

<!-- ═══ 분석중 ═══ -->
<div class="empty-state" id="loadingState" style="display:none">
<div class="sp on" style="width:40px;height:40px;border-width:4px"></div>
<div class="es-title">데이터 분석중...</div>
<div class="es-desc">공사관리대장을 분석하고 있습니다.<br>잠시만 기다려주세요.</div>
</div>

<!-- ═══ 대시보드 본문 ═══ -->
<div id="dashboardContent">
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
<div class="cd c6"><div class="lb">공사건수</div><div class="vl" id="capCnt">-</div><div class="sb">&nbsp;</div></div>
</div>
<div class="early-exec" id="capEarlyExec">
<div class="ee-header">
<div class="ee-title">&#9889; 투자비 조기집행 목표</div>
<div class="ee-controls">
<label>목표</label><input type="number" id="capTargetPct" value="60" min="0" max="100" onchange="renderEarlyExec()">
<label>%</label>
<label style="margin-left:8px">기한</label>
<select id="capTargetMonth" onchange="renderEarlyExec()">
<option value="3">3월</option><option value="4">4월</option><option value="5">5월</option>
<option value="6" selected>6월</option><option value="7">7월</option><option value="8">8월</option><option value="9">9월</option>
</select>
</div>
</div>
<div id="capEarlyBody"></div>
</div>
<div class="stabs">
<button class="st on" onclick="subTab('cap',this,'capBudget')">예산현황</button>
<button class="st" onclick="subTab('cap',this,'capProj')">공사목록</button>
</div>
<div class="sp2 on" id="capBudgetP">
<div class="cbox"><h3>자본 예산 집행 현황 (백만원)</h3><canvas id="chCapBar"></canvas></div>
<div class="tb"><h3 style="display:flex;justify-content:space-between;align-items:center">자본 예산현황<span style="font-size:11px;font-weight:400;color:var(--text3)">(단위: 원, 배정예산 직접 입력)</span></h3><div class="sc">
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
<div style="display:flex;gap:8px;margin-top:10px">
<button class="add-row-btn" onclick="addRow('cap')">+ 항목 추가</button>
<button class="add-row-btn" style="border-color:var(--navy2);background:var(--navy2);color:#fff" onclick="saveBudgets('cap')">배정예산 저장</button>
</div>
</div>
</div>
<div class="sp2" id="capProjP">
<div class="sf"><input type="text" id="capSrch" placeholder="공사번호/업체명/과목 검색..."><select id="capFilt"><option value="">전체</option><option value="공사완료">완료</option><option value="공사중">진행</option><option value="공사중지">중지</option></select></div>
<div class="tb"><h3>공사목록 (자본)</h3><div class="sc"><table id="tCapProj"><thead><tr><th>No</th><th>공사번호</th><th>자본예산과목</th><th>공사업체</th><th>공사상태</th><th>착공일</th><th>현장시공<br>완료일</th><th>준공일</th><th>설계(자본)</th><th>기성(자본)</th><th>예정(자본)</th><th>기성율</th></tr></thead><tbody></tbody></table></div></div>
</div>
<div class="ai-panel" id="capAiPanel" style="display:none">
<div class="ai-header"><div class="ai-icon">AI</div> AI 예산 분석</div>
<div class="ai-summary" id="capAiSummary"></div>
<div class="ai-alerts" id="capAiAlerts"></div>
</div>
<!-- Delay Risk -->
<!-- What-if -->
<div class="ai-panel whatif-panel" id="capWhatifPanel" style="display:none">
<div class="ai-header"><div class="ai-icon" style="background:linear-gradient(135deg,#10B981,#059669)">&#9889;</div> What-if 시뮬레이션</div>
<div class="whatif-controls">
<label>집행 속도 조정</label>
<input type="range" id="capWhatifSlider" min="50" max="150" value="100" step="5" oninput="updateWhatif('cap')">
<span id="capWhatifPctLabel" class="whatif-pct-label">100%</span>
</div>
<div class="ai-summary" id="capWhatifSummary"></div>
<div class="whatif-chart-wrap"><canvas id="chCapWhatif" style="max-height:200px"></canvas></div>
<div id="capWhatifMessage" class="whatif-message"></div>
</div>
<div class="ai-panel realloc-panel" id="capReallocPanel" style="display:none">
<div class="ai-header"><div class="ai-icon" style="background:linear-gradient(135deg,#7C3AED,#6D28D9)">&#8644;</div> 예산 재배분 추천 (자본)</div>
<div class="ai-summary" id="capReallocSummary"></div>
<div id="capReallocBody"></div>
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
<div class="cd c6"><div class="lb">공사건수</div><div class="vl" id="revCnt">-</div><div class="sb">&nbsp;</div></div>
</div>
<div class="stabs">
<button class="st on" onclick="subTab('rev',this,'revBudget')">예산현황</button>
<button class="st" onclick="subTab('rev',this,'revProj')">공사목록</button>
</div>
<div class="sp2 on" id="revBudgetP">
<div class="cbox"><h3>손익 예산 집행 현황 (백만원)</h3><canvas id="chRevBar"></canvas></div>
<div class="tb"><h3 style="display:flex;justify-content:space-between;align-items:center">손익 예산현황<span style="font-size:11px;font-weight:400;color:var(--text3)">(단위: 원, 배정예산 직접 입력)</span></h3><div class="sc">
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
<div style="display:flex;gap:8px;margin-top:10px">
<button class="add-row-btn" onclick="addRow('rev')">+ 항목 추가</button>
<button class="add-row-btn" style="border-color:var(--navy2);background:var(--navy2);color:#fff" onclick="saveBudgets('rev')">배정예산 저장</button>
</div>
</div>
</div>
<div class="sp2" id="revProjP">
<div class="sf"><input type="text" id="revSrch" placeholder="공사번호/업체명/과목 검색..."><select id="revFilt"><option value="">전체</option><option value="공사완료">완료</option><option value="공사중">진행</option><option value="공사중지">중지</option></select></div>
<div class="tb"><h3>공사목록 (손익)</h3><div class="sc"><table id="tRevProj"><thead><tr><th>No</th><th>공사번호</th><th>손익예산과목</th><th>공사업체</th><th>공사상태</th><th>착공일</th><th>현장시공<br>완료일</th><th>준공일</th><th>설계(손익)</th><th>기성(손익)</th><th>예정(손익)</th><th>기성율</th></tr></thead><tbody></tbody></table></div></div>
</div>
<div class="ai-panel" id="revAiPanel" style="display:none">
<div class="ai-header"><div class="ai-icon">AI</div> AI 예산 분석</div>
<div class="ai-summary" id="revAiSummary"></div>
<div class="ai-alerts" id="revAiAlerts"></div>
</div>
<!-- Delay Risk -->
<!-- What-if -->
<div class="ai-panel whatif-panel" id="revWhatifPanel" style="display:none">
<div class="ai-header"><div class="ai-icon" style="background:linear-gradient(135deg,#10B981,#059669)">&#9889;</div> What-if 시뮬레이션</div>
<div class="whatif-controls">
<label>집행 속도 조정</label>
<input type="range" id="revWhatifSlider" min="50" max="150" value="100" step="5" oninput="updateWhatif('rev')">
<span id="revWhatifPctLabel" class="whatif-pct-label">100%</span>
</div>
<div class="ai-summary" id="revWhatifSummary"></div>
<div class="whatif-chart-wrap"><canvas id="chRevWhatif" style="max-height:200px"></canvas></div>
<div id="revWhatifMessage" class="whatif-message"></div>
</div>
<div class="ai-panel realloc-panel" id="revReallocPanel" style="display:none">
<div class="ai-header"><div class="ai-icon" style="background:linear-gradient(135deg,#7C3AED,#6D28D9)">&#8644;</div> 예산 재배분 추천 (손익)</div>
<div class="ai-summary" id="revReallocSummary"></div>
<div id="revReallocBody"></div>
</div>
</div>
</div><!-- /dashboardContent -->

<!-- ═══ 보고서 모달 ═══ -->
<div class="modal-overlay" id="reportModal">
<div class="modal-box">
<div class="modal-head">
<h2>예산 집행 현황 보고서</h2>
<div class="modal-actions">
<button class="btn-copy" onclick="copyReport()">복사</button>
<button class="btn-close" onclick="closeReport()">닫기</button>
</div>
</div>
<div class="modal-body"><pre id="reportText"></pre></div>
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
    document.getElementById('loadingState').style.display='flex';
    document.getElementById('dashboardContent').style.display='none';
    // 기존 배정예산 보존
    const savedBudgets={};
    if(D){
        ['capital','revenue'].forEach(dk=>{
            savedBudgets[dk]={};
            D[dk].budget_comparison.forEach(r=>{
                if(r['배정예산']>0) savedBudgets[dk][r['예산과목']]=r['배정예산'];
            });
        });
    }
    const fd=new FormData();fd.append('file',f);
    try{const r=await fetch('/api/analyze',{method:'POST',body:fd});const j=await r.json();if(j.error){alert(j.error);document.getElementById('loadingState').style.display='none';document.getElementById('dashboardContent').style.display='block';return}D=j;
        // 배정예산 복원
        ['capital','revenue'].forEach(dk=>{
            const map=savedBudgets[dk]||{};
            D[dk].budget_comparison.forEach(r=>{
                const name=r['예산과목'];
                if(name in map){
                    const a=map[name];
                    r['배정예산']=a;r['배정예산_백만']=a/1e6;
                    const d=r['집행실적'],g=r['예상집행'];
                    r['잔액']=a-d;r['예상잔액']=a-g;
                    r['집행율']=a?+(d/a*100).toFixed(1):0;
                    r['예상집행율']=a?+(g/a*100).toFixed(1):0;
                    r['상태']=r['예상집행율']>100?'초과':r['예상집행율']>70?'양호':'미달';
                }
            });
        });
        document.getElementById('loadingState').style.display='none';document.getElementById('dashboardContent').style.display='block';document.getElementById('refreshBtn').style.display='inline-block';document.getElementById('reportBtn').style.display='inline-block';renderAll();updateSummaryCards('cap');updateSummaryCards('rev')}
    catch(err){alert(err.message);document.getElementById('loadingState').style.display='none';document.getElementById('dashboardContent').style.display='block'}
});

async function doRefresh(){
    // 업로드 데이터 + 배정예산 전체 초기화
    try{await fetch('/api/reset');}catch(e){}
    document.getElementById('refreshBtn').style.display='none';
    document.getElementById('reportBtn').style.display='none';
    document.getElementById('fn').textContent='';
    ['cap','rev'].forEach(p=>{
        ['AiPanel','WhatifPanel','ReallocPanel'].forEach(s=>{
            const el=document.getElementById(p+s);if(el)el.style.display='none';
        });
    });
    try{const r=await fetch('/api/init');const j=await r.json();D=j;renderAll();updateSummaryCards('cap');updateSummaryCards('rev');}catch(e){}
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

function renderAll(){if(!D)return;renderCapital();renderRevenue();renderProjects();renderAI();renderEarlyExec();renderWhatif();renderReallocation()}

// ═══════════════════════════════════════
// 투자비 조기집행 분석
// ═══════════════════════════════════════
function renderEarlyExec(){
    if(!D)return;
    const comp=D.capital.budget_comparison;
    const pct=parseInt(document.getElementById('capTargetPct').value)||60;
    const deadline=parseInt(document.getElementById('capTargetMonth').value)||6;
    const now=new Date();
    const curMonth=now.getMonth()+1;
    const totalBudget=comp.reduce((s,r)=>s+r['배정예산'],0);
    const totalExec=comp.reduce((s,r)=>s+r['집행실적'],0);
    const totalForecast=comp.reduce((s,r)=>s+r['예상집행'],0);
    const targetAmt=totalBudget*pct/100;
    const gap=targetAmt-totalExec;
    const achievedPct=targetAmt>0?+(totalExec/targetAmt*100).toFixed(1):0;
    const isOverdue=curMonth>deadline;
    const remainMonths=Math.max(0,deadline-curMonth);

    const body=document.getElementById('capEarlyBody');
    if(totalBudget<=0){body.innerHTML='<div style="font-size:12px;color:#92400E;text-align:center;padding:8px">배정예산을 먼저 입력하세요</div>';return}

    let html='';
    // 전체 요약
    const barPct=Math.min(100,achievedPct);
    const barColor=achievedPct>=100?'var(--green)':achievedPct>=70?'var(--amber)':'var(--red)';
    const statusText=achievedPct>=100?'목표 달성':isOverdue?'기한 초과':'진행중';
    const statusColor=achievedPct>=100?'var(--green)':isOverdue?'var(--red)':'var(--amber)';

    html+=`<div class="ee-total">`;
    html+=`<div><div class="ee-total-label">목표: 전체 배정예산의 ${pct}% (${deadline}월 이내)</div>`;
    html+=`<div style="display:flex;align-items:center;gap:12px;margin-top:6px">`;
    html+=`<div class="ee-total-value">${(targetAmt/1e8).toFixed(1)}억 중 ${(totalExec/1e8).toFixed(1)}억 집행</div>`;
    html+=`<span class="bg" style="background:${statusColor};color:#fff;padding:3px 10px;font-size:10px">${statusText}</span>`;
    html+=`</div></div>`;
    html+=`<div style="text-align:right"><div class="ee-total-value" style="color:${barColor}">${achievedPct}%</div>`;
    html+=`<div class="ee-total-sub">${gap>0?'부족: '+(gap/1e8).toFixed(1)+'억':''}</div></div>`;
    html+=`</div>`;

    // 전체 프로그레스바
    html+=`<div style="margin-top:8px;height:12px;background:#F3F4F6;border-radius:6px;overflow:hidden;position:relative">`;
    html+=`<div style="height:100%;width:${barPct}%;background:${barColor};border-radius:6px;transition:width .3s"></div>`;
    html+=`</div>`;

    // 남은 기간 안내
    if(gap>0&&!isOverdue&&remainMonths>0){
        const monthlyNeeded=gap/remainMonths;
        html+=`<div style="margin-top:10px;padding:8px 12px;background:#FEF3C7;border-radius:6px;font-size:11px;color:#92400E">`;
        html+=`&#128197; ${deadline}월까지 ${remainMonths}개월 남음 &mdash; 월평균 <b>${(monthlyNeeded/1e8).toFixed(2)}억</b> 추가 집행 필요`;
        html+=`</div>`;
    }else if(gap>0&&isOverdue){
        html+=`<div style="margin-top:10px;padding:8px 12px;background:#FEE2E2;border-radius:6px;font-size:11px;color:#991B1B">`;
        html+=`&#9888; 기한(${deadline}월) 경과 &mdash; 목표 대비 <b>${(gap/1e8).toFixed(1)}억</b> 미달`;
        html+=`</div>`;
    }else if(gap<=0){
        html+=`<div style="margin-top:10px;padding:8px 12px;background:#D1FAE5;border-radius:6px;font-size:11px;color:#065F46">`;
        html+=`&#10003; 조기집행 목표 달성! 배정예산의 ${pct}% 이상 집행 완료`;
        html+=`</div>`;
    }

    body.innerHTML=html;
}

// ═══════════════════════════════════════
// AI 분석 패널 렌더
// ═══════════════════════════════════════
function renderAI(){
    if(!D||!D.ai_analysis)return;
    renderAIPanel('cap','capital');
    renderAIPanel('rev','revenue');
}

function renderAIPanel(prefix,dataKey){
    const ai=D.ai_analysis[dataKey];
    const panel=document.getElementById(prefix+'AiPanel');
    if(!ai||!ai.predictions||Array.isArray(ai.predictions)||!ai.predictions.elapsed_pct){
        if(panel)panel.style.display='none';
        return;
    }
    const summaryEl=document.getElementById(prefix+'AiSummary');
    const alertsEl=document.getElementById(prefix+'AiAlerts');
    const pred=ai.predictions;
    const anomalies=ai.anomalies;
    document.getElementById('reportBtn').style.display='inline-block';

    // Summary stats
    let html='';
    html+=`<div class="ai-stat"><div class="ai-label">연간 경과율</div><div class="ai-value">${pred.elapsed_pct}%</div><div class="ai-sub">${pred.month}월 / 12월</div></div>`;
    html+=`<div class="ai-stat"><div class="ai-label">연말 예측 집행액</div><div class="ai-value">${(pred.total_predicted/1e8).toFixed(1)}억</div><div class="ai-sub">예측 집행율 ${pred.predicted_rate}%</div></div>`;
    html+=`<div class="ai-stat"><div class="ai-label">예측 신뢰도</div><div class="ai-value">${pred.confidence}</div><div class="ai-sub">${pred.month>=9?'데이터 충분':'추가 데이터 필요'}</div></div>`;

    // Risk items count
    const riskItems=pred.items.filter(i=>i['위험도']!=='정상');
    html+=`<div class="ai-stat"><div class="ai-label">위험/주의 항목</div><div class="ai-value" style="color:${riskItems.length>0?'var(--red)':'var(--green)'}">${riskItems.length}건</div><div class="ai-sub">${riskItems.length>0?riskItems.map(i=>i['예산과목'].substring(0,6)).join(', '):'이상 없음'}</div></div>`;
    summaryEl.innerHTML=html;

    // Alerts
    let alertHtml='';
    if(anomalies.length===0){
        alertHtml='<div class="ai-alert info"><div class="ai-alert-icon">&#10003;</div><div class="ai-alert-text"><div class="ai-alert-title">이상 항목 없음</div><div class="ai-alert-detail">현재 예산 집행이 정상 범위 내에 있습니다.</div></div></div>';
    }else{
        const icons={danger:'!',warning:'!',info:'i'};
        anomalies.slice(0,5).forEach(a=>{
            alertHtml+=`<div class="ai-alert ${a.level}"><div class="ai-alert-icon">${icons[a.level]||'i'}</div><div class="ai-alert-text"><div class="ai-alert-title">${a.category}: ${a.message}</div><div class="ai-alert-detail">${a.detail}</div></div></div>`;
        });
        if(anomalies.length>5){
            alertHtml+=`<div style="font-size:10px;color:var(--text3);text-align:center;padding:4px">외 ${anomalies.length-5}건</div>`;
        }
    }
    alertsEl.innerHTML=alertHtml;
    panel.style.display='block';
}

// ═══════════════════════════════════════
// 보고서 모달
// ═══════════════════════════════════════
function openReport(){
    if(!D||!D.ai_analysis)return;
    document.getElementById('reportText').textContent=D.ai_analysis.report;
    document.getElementById('reportModal').classList.add('on');
}
function closeReport(){document.getElementById('reportModal').classList.remove('on')}
function copyReport(){
    const text=document.getElementById('reportText').textContent;
    navigator.clipboard.writeText(text).then(()=>{
        const btn=document.querySelector('.btn-copy');
        btn.textContent='복사됨!';
        setTimeout(()=>{btn.textContent='복사'},1500);
    });
}
document.getElementById('reportModal').addEventListener('click',function(e){if(e.target===this)closeReport()});

// ═══════════════════════════════════════
// Feature 1: 예산 소진 예측 곡선
// ═══════════════════════════════════════

// ═══════════════════════════════════════

// ═══════════════════════════════════════
// Feature 3: What-if 시뮬레이션
// ═══════════════════════════════════════
function renderWhatif(){
    if(!D||!D.ai_analysis)return;
    _initWhatif('cap','capital');
    _initWhatif('rev','revenue');
}
function _initWhatif(prefix,dataKey){
    const ai=D.ai_analysis[dataKey];
    if(!ai||!ai.whatif)return;
    if(ai.whatif.total_budget<=0){document.getElementById(prefix+'WhatifPanel').style.display='none';return}
    document.getElementById(prefix+'WhatifPanel').style.display='block';
    document.getElementById(prefix+'WhatifSlider').value=100;
    document.getElementById(prefix+'WhatifPctLabel').textContent='100%';
    updateWhatif(prefix);
}
function updateWhatif(prefix){
    const dataKey=prefix==='cap'?'capital':'revenue';
    const ai=D.ai_analysis[dataKey];
    if(!ai||!ai.whatif)return;
    const wi=ai.whatif;
    const slider=document.getElementById(prefix+'WhatifSlider');
    const pctLabel=document.getElementById(prefix+'WhatifPctLabel');
    const summaryEl=document.getElementById(prefix+'WhatifSummary');
    const msgEl=document.getElementById(prefix+'WhatifMessage');
    const factor=parseInt(slider.value)/100;
    pctLabel.textContent=slider.value+'%';
    const adjustedRate=wi.monthly_rate*factor;
    const remainExec=adjustedRate*wi.remaining_months;
    const projYE=wi.total_exec+remainExec;
    const projRate=wi.total_budget>0?+(projYE/wi.total_budget*100).toFixed(1):0;
    const need90=(wi.target_90-wi.total_exec)/Math.max(1,wi.remaining_months);
    const need100=(wi.total_budget-wi.total_exec)/Math.max(1,wi.remaining_months);
    let html='';
    html+='<div class="ai-stat"><div class="ai-label">조정 월집행액</div><div class="ai-value">'+(adjustedRate/1e8).toFixed(2)+'억</div><div class="ai-sub">기존 '+(wi.monthly_rate/1e8).toFixed(2)+'억 x '+slider.value+'%</div></div>';
    html+='<div class="ai-stat"><div class="ai-label">예상 연말 집행</div><div class="ai-value" style="color:'+(projRate>100?'var(--red)':projRate>90?'var(--green)':'#F59E0B')+'">'+(projYE/1e8).toFixed(1)+'억</div><div class="ai-sub">예상집행율 '+projRate+'%</div></div>';
    html+='<div class="ai-stat"><div class="ai-label">90% 달성 필요</div><div class="ai-value">'+(wi.remaining_months>0?(need90/1e8).toFixed(2)+'억/월':'-')+'</div><div class="ai-sub">'+wi.remaining_months+'개월 남음</div></div>';
    html+='<div class="ai-stat"><div class="ai-label">100% 달성 필요</div><div class="ai-value">'+(wi.remaining_months>0?(need100/1e8).toFixed(2)+'억/월':'-')+'</div><div class="ai-sub">배정예산 완전소진</div></div>';
    summaryEl.innerHTML=html;
    // 미니 차트
    const lb=[];const curL=[];const adjL=[];const bL=[];
    for(let m=1;m<=12;m++){
        lb.push(m+'월');
        if(m<=wi.current_month){
            const v=+(wi.total_exec*m/wi.current_month/1e8).toFixed(2);
            curL.push(v);adjL.push(v);
        }else{
            curL.push(+((wi.total_exec+wi.monthly_rate*(m-wi.current_month))/1e8).toFixed(2));
            adjL.push(+((wi.total_exec+adjustedRate*(m-wi.current_month))/1e8).toFixed(2));
        }
        bL.push(+(wi.total_budget/1e8).toFixed(2));
    }
    const cid=prefix==='cap'?'chCapWhatif':'chRevWhatif';
    _ch(cid,'line',{labels:lb,datasets:[
        {label:'배정예산',data:bL,borderColor:'rgba(59,130,246,0.5)',borderDash:[6,3],borderWidth:1.5,pointRadius:0,fill:false},
        {label:'현재페이스',data:curL,borderColor:'rgba(148,163,184,0.6)',borderWidth:1.5,borderDash:[4,2],pointRadius:0,fill:false},
        {label:'조정 후',data:adjL,borderColor:'rgba(16,185,129,1)',borderWidth:2.5,pointRadius:2,pointBackgroundColor:'rgba(16,185,129,1)',fill:false}
    ]},{animation:{duration:0},scales:{y:{beginAtZero:true,title:{display:true,text:'억원'},ticks:{font:{size:9}}},x:{ticks:{font:{size:9}}}},
        plugins:{legend:{position:'bottom',labels:{font:{size:9},usePointStyle:true}}}});
    // 메시지
    let msg='';let msgCls='';
    if(projRate>=95&&projRate<=105){msg='현재 속도의 '+slider.value+'%로 집행 시, 연말 집행율 '+projRate+'% 예상. 적정 수준입니다.';msgCls='whatif-msg-ok';}
    else if(projRate>105){msg='예산 초과 위험! 연말 '+projRate+'% 집행 예상. 속도를 줄여야 합니다.';msgCls='whatif-msg-danger';}
    else if(projRate>=80){msg='연말 집행율 '+projRate+'% 예상. 목표 달성을 위해 속도 조정을 검토하세요.';msgCls='whatif-msg-warn';}
    else{msg='연말 집행율 '+projRate+'% 예상. 집행 부진 우려. 하반기 집중 집행이 필요합니다.';msgCls='whatif-msg-danger';}
    if(prefix==='cap'){
        const earlyPct=parseInt(document.getElementById('capTargetPct').value)||60;
        const earlyMonth=parseInt(document.getElementById('capTargetMonth').value)||6;
        const earlyTarget=wi.total_budget*earlyPct/100;
        if(wi.total_exec<earlyTarget&&wi.current_month<=earlyMonth){
            const mToT=earlyMonth-wi.current_month;
            const needPM=(earlyTarget-wi.total_exec)/Math.max(1,mToT);
            msg+=adjustedRate>=needPM?' | 조기집행 '+earlyPct+'% 달성 가능':' | 조기집행 '+earlyPct+'% 미달 예상 (월 '+(needPM/1e8).toFixed(2)+'억 필요)';
        }
    }
    msgEl.className='whatif-message '+msgCls;
    msgEl.innerHTML=msg;
}

// ═══════════════════════════════════════
// Feature 6: 예산 재배분 추천
// ═══════════════════════════════════════
function renderReallocation(){
    if(!D||!D.ai_analysis)return;
    _renderReallocPanel('cap','capital');
    _renderReallocPanel('rev','revenue');
}
function _renderReallocPanel(prefix,dataKey){
    const ai=D.ai_analysis[dataKey];
    if(!ai||!ai.reallocation)return;
    const ra=ai.reallocation;
    const panel=document.getElementById(prefix+'ReallocPanel');
    const summaryEl=document.getElementById(prefix+'ReallocSummary');
    const bodyEl=document.getElementById(prefix+'ReallocBody');
    if(!ra.recommendations||ra.recommendations.length===0){
        const hasBudget=D[dataKey].budget_comparison.some(r=>r['배정예산']>0);
        if(!hasBudget){panel.style.display='none';return}
        panel.style.display='block';
        summaryEl.innerHTML='<div class="ai-stat" style="grid-column:1/-1"><div class="ai-label">분석 결과</div><div class="ai-value" style="color:var(--green)">재배분 불필요</div><div class="ai-sub">모든 항목이 적정 범위 내에서 집행 중입니다</div></div>';
        bodyEl.innerHTML='';
        return;
    }
    panel.style.display='block';
    let html='';
    html+='<div class="ai-stat"><div class="ai-label">재배분 추천</div><div class="ai-value">'+ra.count+'건</div><div class="ai-sub">잉여 → 부족 이전</div></div>';
    html+='<div class="ai-stat"><div class="ai-label">이전 가능 총액</div><div class="ai-value">'+(ra.total_transferable/1e8).toFixed(1)+'억</div><div class="ai-sub">추천 이전 합계</div></div>';
    summaryEl.innerHTML=html;
    let bodyHtml='';
    ra.recommendations.forEach(r=>{
        bodyHtml+='<details class="realloc-details">';
        bodyHtml+='<summary>';
        bodyHtml+='<span class="realloc-toggle">&#9658;</span>';
        bodyHtml+='<span class="rc-box" style="flex:1"><span class="rc-name">'+r.source+' &#10132; '+r.target+'</span><span class="rc-detail" style="display:inline-block;margin-left:6px;color:var(--text3)">이전 가능액</span></span>';
        bodyHtml+='<span class="rc-amount">'+(r.amount/1e8).toFixed(2)+'억</span>';
        bodyHtml+='</summary>';
        bodyHtml+='<div class="realloc-details-body">';
        bodyHtml+='<div class="rc-box"><div class="rc-name">'+r.source+'</div><div class="rc-detail">예산 '+(r.source_budget/1e8).toFixed(1)+'억 | 예상집행율 <span class="rc-rate-change" style="color:var(--green)">'+r.source_fc_rate+'%</span></div></div>';
        bodyHtml+='<div class="rc-arrow">&#10132;</div>';
        bodyHtml+='<div class="rc-amount">'+(r.amount/1e8).toFixed(2)+'억</div>';
        bodyHtml+='<div class="rc-arrow">&#10132;</div>';
        bodyHtml+='<div class="rc-box"><div class="rc-name">'+r.target+'</div><div class="rc-detail">예산 '+(r.target_budget/1e8).toFixed(1)+'억 | 예상집행율 <span class="rc-rate-change" style="color:var(--red)">'+r.target_fc_rate+'%</span> &#8594; <span class="rc-rate-change" style="color:var(--green)">'+r.target_new_rate+'%</span></div></div>';
        bodyHtml+='<div class="rc-reason">'+r.reason+'</div>';
        bodyHtml+='</div>';
        bodyHtml+='</details>';
    });
    bodyEl.innerHTML=bodyHtml;
}

// ═══════════════════════════════════════
// 차트 렌더 (분리)
// ═══════════════════════════════════════
function _renderBarChart(chartId,comp){
    const lb=[],aV=[],dV=[],gV=[];
    comp.forEach(r=>{if(!r['배정예산_백만']&&!r['집행실적']&&!r['예상집행'])return;lb.push(sn(r['예산과목'],7));aV.push(r['배정예산_백만']);dV.push(Math.round(r['집행실적']/1e6));gV.push(Math.round(r['예상집행']/1e6))});
    const mx=aV.length?Math.max(...aV,...dV,...gV):0;const yMax=niceMax(mx);
    _ch(chartId,'bar',{labels:lb,datasets:[
        {label:'배정예산(A)',data:aV,backgroundColor:'rgba(59,130,246,.7)',borderColor:'rgba(59,130,246,1)',borderWidth:1,borderRadius:3},
        {label:'집행실적(D)',data:dV,backgroundColor:'rgba(239,68,68,.65)',borderColor:'rgba(239,68,68,1)',borderWidth:1,borderRadius:3},
        {label:'최종예상(G)',data:gV,backgroundColor:'rgba(251,191,36,.6)',borderColor:'rgba(245,158,11,1)',borderWidth:1,borderRadius:3}
    ]},{scales:{y:{beginAtZero:true,max:yMax,afterBuildTicks(axis){const t=[0,50,100];for(let v=200;v<=axis.max;v+=200)t.push(v);axis.ticks=t.map(v=>({value:v}))},ticks:{font:{size:10}},title:{display:true,text:'백만원'}}}});
}
function renderCapitalChart(){_renderBarChart('chCapBar',D.capital.budget_comparison)}
function renderRevenueChart(){_renderBarChart('chRevBar',D.revenue.budget_comparison)}

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
    renderCapitalChart();

    const comp=D.capital.budget_comparison;
    const tb=document.querySelector('#tCapMain tbody');tb.innerHTML='';
    comp.forEach((r,i)=>{
        const a=r['배정예산'],b=r['소비금액'],c=r['약정금액'],d=r['집행실적'],e=r['잔액'],f=r['진행중공사비'],g=r['예상집행'],ga=r['예상잔액'],dr=r['집행율'],gr=r['예상집행율'];
        const tr=document.createElement('tr');
        const codeCell=`<td class="dbl-code" ondblclick="startEditCode(this,'capital',${i})">${r['사업코드']||''}</td>`;
        const nameCell=r._custom
            ?`<td><input class="inp-code" type="text" value="${r['예산과목']}" placeholder="사업명" onchange="D.capital.budget_comparison[${i}]['예산과목']=this.value"></td>`
            :`<td>${r['예산과목']}</td>`;
        tr.innerHTML=codeCell+nameCell+`
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
    renderRevenueChart();

    const comp=D.revenue.budget_comparison;
    const tb=document.querySelector('#tRevMain tbody');tb.innerHTML='';
    comp.forEach((r,i)=>{
        const a=r['배정예산'],b=r['소비금액'],c=r['약정금액'],d=r['집행실적'],e=r['잔액'],f=r['진행중공사비'],g=r['예상집행'],ga=r['예상잔액'],dr=r['집행율'],gr=r['예상집행율'];
        const tr=document.createElement('tr');
        const codeCell=`<td class="dbl-code" ondblclick="startEditCode(this,'revenue',${i})">${r['사업코드']||''}</td>`;
        const nameCell=r._custom
            ?`<td><input class="inp-code" type="text" value="${r['예산과목']}" placeholder="사업명" onchange="D.revenue.budget_comparison[${i}]['예산과목']=this.value"></td>`
            :`<td>${r['예산과목']}</td>`;
        tr.innerHTML=codeCell+nameCell+`
            <td><input class="inp-budget" type="text" data-section="rev" data-idx="${i}" value="${fw(a)}" onchange="recalcBudget(this)"></td>
            <td>${fw(b)}</td><td>${fw(c)}</td><td>${fw(d)}</td>
            <td class="${clr(e)}">${fw(e)}</td><td>${br(dr)}</td>
            <td>${fw(f)}</td><td>${fw(g)}</td><td class="${clr(ga)}">${fw(ga)}</td><td>${br(gr)}</td><td>${r['건수']}</td>`;
        tb.appendChild(tr);
    });
    updateTotals('rev');
}

// ═══════════════════════════════════════
// 사업코드 더블클릭 인라인 편집
// ═══════════════════════════════════════
function startEditCode(td, dataKey, idx){
    const prev=td.textContent.trim();
    td.classList.remove('dbl-code');
    td.ondblclick=null;
    const inp=document.createElement('input');
    inp.className='inp-code';
    inp.value=prev;
    inp.placeholder='사업코드';
    td.textContent='';
    td.appendChild(inp);
    inp.focus();inp.select();
    function commit(){
        const val=inp.value.trim();
        D[dataKey].budget_comparison[idx]['사업코드']=val;
        td.textContent=val;
        td.classList.add('dbl-code');
        td.ondblclick=()=>startEditCode(td,dataKey,idx);
    }
    inp.addEventListener('blur',commit);
    inp.addEventListener('keydown',e=>{if(e.key==='Enter'){inp.blur();}if(e.key==='Escape'){inp.value=prev;inp.blur();}});
}

// ═══════════════════════════════════════
// 배정예산 저장 버튼
// ═══════════════════════════════════════
function saveBudgets(sec){
    const dataKey=sec==='cap'?'capital':'revenue';
    const tId=sec==='cap'?'tCapMain':'tRevMain';
    // 테이블의 모든 배정예산 input 읽기
    document.querySelectorAll('#'+tId+' .inp-budget').forEach(inp=>{
        const idx=parseInt(inp.dataset.idx);
        const val=parseFloat(inp.value.replace(/,/g,''))||0;
        const r=D[dataKey].budget_comparison[idx];
        if(!r)return;
        r['배정예산']=val;r['배정예산_백만']=val/1e6;
        const d=r['집행실적'],g=r['예상집행'];
        r['잔액']=val-d;r['예상잔액']=val-g;
        r['집행율']=val?+(d/val*100).toFixed(1):0;
        r['예상집행율']=val?+(g/val*100).toFixed(1):0;
        r['상태']=r['예상집행율']>100?'초과':r['예상집행율']>70?'양호':'미달';
        inp.value=fw(val);
        // 행 셀 업데이트
        const tr=inp.closest('tr');const cells=tr.querySelectorAll('td');
        cells[6].className=clr(r['잔액']);cells[6].textContent=fw(r['잔액']);
        cells[7].innerHTML=br(r['집행율']);
        cells[10].className=clr(r['예상잔액']);cells[10].textContent=fw(r['예상잔액']);
        cells[11].innerHTML=br(r['예상집행율']);
    });
    updateTotals(sec);
    updateSummaryCards(sec);
    if(sec==='cap'){renderEarlyExec();renderCapitalChart();renderWhatif();renderReallocation()}
    else{renderRevenueChart();renderWhatif();renderReallocation()}
    // 서버에 저장 (기본항목 금액 + 사용자추가 항목 전체)
    const saveData={capital:{budgets:{},custom_items:[]},revenue:{budgets:{},custom_items:[]}};
    ['capital','revenue'].forEach(k=>{
        saveData[k].codes={};
        (D[k].budget_comparison||[]).forEach(r=>{
            if(r._custom){
                if(r['예산과목']) saveData[k].custom_items.push({사업코드:r['사업코드'],예산과목:r['예산과목'],배정예산:r['배정예산']||0});
            } else {
                if(r['배정예산']>0) saveData[k].budgets[r['예산과목']]=r['배정예산'];
                if(r['사업코드']) saveData[k].codes[r['예산과목']]=r['사업코드'];
            }
        });
    });
    const btn=event.target;
    fetch('/api/save-budgets',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(saveData)})
    .then(r=>r.json()).then(j=>{
        if(j.ok){btn.textContent='저장됨!';btn.style.background='var(--green)';btn.style.borderColor='var(--green)';}
        else{btn.textContent='저장실패';btn.style.background='#e74c3c';btn.style.borderColor='#e74c3c';}
        setTimeout(()=>{btn.textContent='배정예산 저장';btn.style.background='var(--navy2)';btn.style.borderColor='var(--navy2)'},1500);
    }).catch(()=>{
        btn.textContent='저장실패';btn.style.background='#e74c3c';btn.style.borderColor='#e74c3c';
        setTimeout(()=>{btn.textContent='배정예산 저장';btn.style.background='var(--navy2)';btn.style.borderColor='var(--navy2)'},1500);
    });
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
    if(sec==='cap')renderEarlyExec();
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
}

// ═══════════════════════════════════════
// 항목 추가
// ═══════════════════════════════════════
function addRow(sec){
    const dataKey=sec==='cap'?'capital':'revenue';
    const comp=D[dataKey].budget_comparison;
    const tId=sec==='cap'?'tCapMain':'tRevMain';
    const idx=comp.length;
    const newItem={'사업코드':'','예산과목':'','배정예산_백만':0,'배정예산':0,'소비금액':0,'약정금액':0,'집행실적':0,'잔액':0,'집행율':0,'진행중공사비':0,'예상집행':0,'예상잔액':0,'예상집행율':0,'건수':0,'상태':'미달','_custom':true};
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
            tr.innerHTML=`<td>${idx+1}</td><td>${r['공사번호']}</td><td>${r[catKey]}</td><td>${r['공사업체']}</td><td>${stBg(r['공사상태'])}</td><td>${r['착공일']}</td><td>${r['현장시공완료일']||''}</td><td>${r['준공일']||''}</td><td>${fw(r[dK])}</td><td>${fw(r[pK])}</td><td>${fw(r[eK])}</td><td>${br(rate)}</td>`;
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
    // 빈 데이터로 대시보드 바로 표시 (배정예산 먼저 입력 가능)
    try{const r=await fetch('/api/init');const j=await r.json();D=j;renderAll()}catch(e){}
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


def _load_saved_budgets():
    """budgets.json에서 저장된 배정예산 로드"""
    if os.path.exists(BUDGET_FILE):
        try:
            with open(BUDGET_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _apply_saved_budgets(comp_list, section_key, saved):
    """comparison 리스트에 저장된 배정예산 + 사용자 추가 항목 적용"""
    sec_data = saved.get(section_key, {})
    # 기본항목 배정예산 복원
    bmap = sec_data.get('budgets', {}) if isinstance(sec_data, dict) else {}
    # 구버전 호환 (flat dict)
    if isinstance(sec_data, dict) and 'budgets' not in sec_data and 'custom_items' not in sec_data:
        bmap = sec_data
    # 사업코드 복원
    codes_map = sec_data.get('codes', {}) if isinstance(sec_data, dict) else {}
    for r in comp_list:
        name = r['예산과목']
        if name in bmap and bmap[name] > 0:
            a = bmap[name]
            r['배정예산'] = a
            r['배정예산_백만'] = a / 1e6
            d, g = r['집행실적'], r['예상집행']
            r['잔액'] = a - d
            r['예상잔액'] = a - g
            r['집행율'] = round(d / a * 100, 1) if a else 0
            r['예상집행율'] = round(g / a * 100, 1) if a else 0
            r['상태'] = '초과' if r['예상집행율'] > 100 else ('양호' if r['예상집행율'] > 70 else '미달')
        if name in codes_map:
            r['사업코드'] = codes_map[name]
    # 사용자 추가 항목 복원
    custom_items = sec_data.get('custom_items', []) if isinstance(sec_data, dict) else []
    existing_names = {r['예산과목'] for r in comp_list}
    for ci in custom_items:
        if ci.get('예산과목') and ci['예산과목'] not in existing_names:
            a = ci.get('배정예산', 0)
            comp_list.append({
                '사업코드': ci.get('사업코드', ''), '예산과목': ci['예산과목'],
                '배정예산_백만': a / 1e6, '배정예산': a,
                '소비금액': 0, '약정금액': 0,
                '집행실적': 0, '잔액': a,
                '집행율': 0, '진행중공사비': 0, '예상집행': 0,
                '예상잔액': a, '예상집행율': 0,
                '건수': 0, '상태': '미달', '_custom': True,
            })


@app.route('/api/init')
def api_init():
    """배정예산 입력용 빈 데이터 구조 반환 (저장된 배정예산 복원)"""
    def _empty_comp(budget_dict):
        return [{
            '사업코드': bcode, '예산과목': bname,
            '배정예산_백만': 0, '배정예산': 0,
            '소비금액': 0, '약정금액': 0,
            '집행실적': 0, '잔액': 0,
            '집행율': 0, '진행중공사비': 0, '예상집행': 0,
            '예상잔액': 0, '예상집행율': 0,
            '건수': 0, '상태': '미달',
        } for bname, (bcode, bval_mil) in budget_dict.items()]

    def _summary(comp):
        ta = sum(r['배정예산'] for r in comp)
        td = sum(r['집행실적'] for r in comp)
        tg = sum(r['예상집행'] for r in comp)
        return {'배정예산': ta, '소비금액': 0, '약정금액': 0,
                '집행실적': td, '잔액': ta - td,
                '집행율': round(td / ta * 100, 1) if ta else 0,
                '진행중공사비': 0, '예상집행': tg, '예상잔액': ta - tg,
                '예상집행율': round(tg / ta * 100, 1) if ta else 0,
                '공사건수': 0, '초과항목': 0}

    cap_comp = _empty_comp(BUDGET_CAPITAL)
    rev_comp = _empty_comp(BUDGET_REVENUE)

    # 저장된 배정예산 복원
    saved = _load_saved_budgets()
    _apply_saved_budgets(cap_comp, 'capital', saved)
    _apply_saved_budgets(rev_comp, 'revenue', saved)

    return jsonify({
        'capital': {'summary': _summary(cap_comp), 'budget_comparison': cap_comp,
                    'category': [], 'status': {}, 'budget_sheets': {}},
        'revenue': {'summary': _summary(rev_comp), 'budget_comparison': rev_comp,
                    'category': [], 'status': {}, 'budget_sheets': {}},
        'projects': [],
        'ai_analysis': {
            'capital': {'predictions': [], 'anomalies': []},
            'revenue': {'predictions': [], 'anomalies': []},
            'report': '',
        },
    })


@app.route('/api/save-budgets', methods=['POST'])
def api_save_budgets():
    """배정예산을 서버에 저장"""
    data = request.get_json()
    try:
        with open(BUDGET_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({'error': '파일명이 비어있습니다.'}), 400
    global LAST_FILE
    filepath = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(filepath)
    LAST_FILE = filepath
    try:
        result = parse_and_analyze(filepath)
        saved = _load_saved_budgets()
        _apply_saved_budgets(result['capital']['budget_comparison'], 'capital', saved)
        _apply_saved_budgets(result['revenue']['budget_comparison'], 'revenue', saved)
        # 배정예산 적용 후 AI 분석 수행
        cap_comp = result['capital']['budget_comparison']
        rev_comp = result['revenue']['budget_comparison']
        result['ai_analysis']['capital']['whatif'] = _whatif_baseline(cap_comp, BUDGET_CAPITAL)
        result['ai_analysis']['revenue']['whatif'] = _whatif_baseline(rev_comp, BUDGET_REVENUE)
        result['ai_analysis']['capital']['reallocation'] = _reallocation_recommendations(cap_comp)
        result['ai_analysis']['revenue']['reallocation'] = _reallocation_recommendations(rev_comp)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': f'분석 오류: {str(e)}'}), 500


@app.route('/api/reset')
def api_reset():
    """업로드 파일 및 배정예산 전체 초기화"""
    global LAST_FILE
    LAST_FILE = None
    # budgets.json 초기화
    try:
        with open(BUDGET_FILE, 'w', encoding='utf-8') as f:
            json.dump({}, f)
    except Exception:
        pass
    return jsonify({'ok': True})


@app.route('/api/refresh')
def api_refresh():
    if not LAST_FILE or not os.path.exists(LAST_FILE):
        return jsonify({'error': '업로드된 파일이 없습니다.'}), 404
    try:
        result = parse_and_analyze(LAST_FILE)
        saved = _load_saved_budgets()
        _apply_saved_budgets(result['capital']['budget_comparison'], 'capital', saved)
        _apply_saved_budgets(result['revenue']['budget_comparison'], 'revenue', saved)
        # 배정예산 적용 후 AI 분석 수행
        cap_comp = result['capital']['budget_comparison']
        rev_comp = result['revenue']['budget_comparison']
        result['ai_analysis']['capital']['whatif'] = _whatif_baseline(cap_comp, BUDGET_CAPITAL)
        result['ai_analysis']['revenue']['whatif'] = _whatif_baseline(rev_comp, BUDGET_REVENUE)
        result['ai_analysis']['capital']['reallocation'] = _reallocation_recommendations(cap_comp)
        result['ai_analysis']['revenue']['reallocation'] = _reallocation_recommendations(rev_comp)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5005)
