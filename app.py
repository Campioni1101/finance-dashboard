import os
from flask import Flask, jsonify, render_template
from data_parser import parse_finance_data
from market_data import get_all_quotes, get_selic_rate

app = Flask(__name__)

_base = os.path.dirname(os.path.abspath(__file__))
# Usa finance_data.xlsx (novo formato) se existir, senão cai no legado
DATA_FILE = (
    os.path.join(_base, 'finance_data.xlsx')
    if os.path.exists(os.path.join(_base, 'finance_data.xlsx'))
    else os.path.join(_base, 'finance.xlsx')
)

META_MENSAL = 0.0095        # 0,95% ao mês
RO_TARGET = 2000.0          # Reserva de Oportunidade alvo
RE_TARGET = 4000.0          # Reserva de Emergência alvo
CRIPTO_TARGET_PCT = 10.0    # % alvo cripto
ETF_TARGET_PCT = 10.0       # % alvo ETF (em construção)


def get_data() -> dict:
    return parse_finance_data(DATA_FILE)


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/api/data', methods=['GET'])
def api_data():
    return jsonify(get_data())


@app.route('/api/market', methods=['GET'])
def api_market():
    quotes = get_all_quotes()
    selic  = get_selic_rate()
    return jsonify({'quotes': quotes, 'selic_anual': selic})


@app.route('/api/insights', methods=['GET'])
def api_insights():
    data = get_data()
    return jsonify(generate_insights(data))


# ── Insight engine ────────────────────────────────────────────────────────────

def generate_insights(data: dict) -> list:
    s = data['summary']
    monthly = data['monthly']
    insights = []

    # 1. Reserva de Emergência
    re = s.get('reserve_emergency') or 0
    if re >= RE_TARGET:
        insights.append({
            'type': 'success', 'icon': '🛡️',
            'title': 'Reserva de Emergência completa',
            'message': f'Meta de R${RE_TARGET:,.2f} atingida.',
            'action': None
        })
    else:
        insights.append({
            'type': 'warning', 'icon': '⚠️',
            'title': 'Reserva de Emergência incompleta',
            'message': f'Atual: R${re:,.2f} | Meta: R${RE_TARGET:,.2f}',
            'action': f'Prioridade máxima: completar R${RE_TARGET - re:,.2f} antes de qualquer aporte variável'
        })

    # 2. Reserva de Oportunidade
    ro = s.get('reserve_opportunity') or 0
    if ro >= RO_TARGET:
        insights.append({
            'type': 'success', 'icon': '✅',
            'title': 'Reserva de Oportunidade completa',
            'message': f'Meta de R${RO_TARGET:,.2f} atingida.',
            'action': None
        })
    else:
        falta = RO_TARGET - ro
        insights.append({
            'type': 'warning', 'icon': '⚡',
            'title': 'Reserva de Oportunidade incompleta',
            'message': f'Atual: R${ro:,.2f} | Meta: R${RO_TARGET:,.2f} | Faltam R${falta:,.2f}',
            'action': f'Ao investir o aporte mensal, reserve R${falta:,.2f} para completar a R.O. primeiro'
        })

    # 3. Cripto vs meta 10%
    total = s.get('total_real') or 0
    crypto = s.get('crypto') or 0
    if total > 0:
        crypto_pct = crypto / total * 100
        diff = CRIPTO_TARGET_PCT - crypto_pct
        if diff > 1.0:
            insights.append({
                'type': 'info', 'icon': '₿',
                'title': f'Cripto abaixo da meta ({crypto_pct:.1f}% de {CRIPTO_TARGET_PCT}%)',
                'message': f'Atual: R${crypto:,.2f} | Para atingir 10%: R${total * CRIPTO_TARGET_PCT / 100:,.2f}',
                'action': f'Aportar ~R${total * diff / 100:,.2f} em cripto para rebalancear'
            })
        elif diff < -1.0:
            insights.append({
                'type': 'neutral', 'icon': '₿',
                'title': f'Cripto acima da meta ({crypto_pct:.1f}%)',
                'message': f'Meta: {CRIPTO_TARGET_PCT}% | Aguardar valorização dos outros ativos',
                'action': 'Não aportar em cripto até reequilibrar a carteira'
            })
        else:
            insights.append({
                'type': 'success', 'icon': '₿',
                'title': f'Cripto na meta ({crypto_pct:.1f}%)',
                'message': f'R${crypto:,.2f} — equilibrado',
                'action': None
            })

    # 4. META vs REAL — últimos 6 meses
    recent = [m for m in monthly if m.get('real') and m.get('meta')][-6:]
    if recent:
        below = [m for m in recent if m['real'] < m['meta']]
        avg_real = sum(m['real'] for m in recent) / len(recent)
        if len(below) >= 4:
            insights.append({
                'type': 'warning', 'icon': '📉',
                'title': f'{len(below)}/6 meses abaixo da META de {META_MENSAL*100:.2f}%',
                'message': f'Média recente: {avg_real*100:.2f}%/mês',
                'action': 'Revise a alocação: aumentar SELIC estabiliza o rendimento'
            })
        elif len(below) <= 1:
            insights.append({
                'type': 'success', 'icon': '📈',
                'title': f'{6-len(below)}/6 meses acima da META',
                'message': f'Média recente: {avg_real*100:.2f}%/mês — acima dos {META_MENSAL*100:.2f}%',
                'action': None
            })

    # 5. Melhor FII por yield acumulado
    alloc_map = {a['sku']: a['value'] for a in s.get('asset_allocation', [])}
    fii_divs = {k: v for k, v in (s.get('fii_dividends') or {}).items() if v}
    fii_yield = {}
    for fii, div in fii_divs.items():
        invested = alloc_map.get(fii)
        if invested and invested > 0:
            fii_yield[fii] = div / invested

    if fii_yield:
        best = max(fii_yield, key=lambda k: fii_yield[k])
        insights.append({
            'type': 'info', 'icon': '🏆',
            'title': f'Melhor FII para reinvestir: {best.upper()}',
            'message': (f'Yield acumulado: {fii_yield[best]*100:.1f}% | '
                        f'Dividendos totais: R${fii_divs[best]:,.2f}'),
            'action': f'Reinvista os dividendos dos FIIs em {best.upper()} para maximizar o yield'
        })

    # 6. Ganho/perda renda variável no último mês
    last_valid = next((m for m in reversed(monthly) if m.get('gain_loss') is not None), None)
    if last_valid:
        gl = last_valid['gain_loss']
        pv = last_valid['portfolio_value'] or 1
        gl_pct = gl / (pv - gl) * 100 if (pv - gl) != 0 else 0
        if gl < 0:
            insights.append({
                'type': 'warning', 'icon': '💼',
                'title': f'R.V. desvalorizada ({last_valid["label"]})',
                'message': (f'Carteira variável: R${pv:,.2f} | '
                            f'Vs. investido: {gl:+,.2f} ({gl_pct:+.1f}%)'),
                'action': 'Renda variável é longo prazo — mantenha posição e acompanhe os fundamentos'
            })
        else:
            insights.append({
                'type': 'success', 'icon': '💼',
                'title': f'R.V. valorizada ({last_valid["label"]})',
                'message': f'Acima do custo em R${gl:,.2f} ({gl_pct:+.1f}%)',
                'action': None
            })

    # 7. Sugestão de onde alocar o aporte mensal
    insights.append({
        'type': 'info', 'icon': '💡',
        'title': 'Sugestão para o próximo aporte',
        'message': _build_contribution_message(s, monthly, alloc_map, fii_yield),
        'action': _build_contribution_action(s, monthly, alloc_map, fii_yield)
    })

    # 8. Onde reinvestir os dividendos
    total_div_last = _last_month_total_dividends(monthly)
    if total_div_last > 0:
        insights.append({
            'type': 'info', 'icon': '🔄',
            'title': f'Dividendos do último mês: R${total_div_last:,.2f}',
            'message': _build_dividend_reinvest_message(s, fii_yield, alloc_map),
            'action': _build_dividend_reinvest_action(fii_yield, alloc_map)
        })

    return insights


def _build_contribution_message(s, monthly, alloc_map, fii_yield) -> str:
    ro = s.get('reserve_opportunity') or 0
    if ro < RO_TARGET:
        return f'R.O. incompleta — prioridade antes de aportes variáveis'
    last_real = next((m['real'] for m in reversed(monthly) if m.get('real')), None)
    if last_real and last_real < META_MENSAL:
        return 'Último mês abaixo da META — SELIC oferece retorno mais estável agora'
    if alloc_map:
        avg = sum(alloc_map.values()) / len(alloc_map)
        underweight = min(alloc_map.items(), key=lambda x: x[1])
        if underweight[1] < avg * 0.85:
            return f'{underweight[0].upper()} está abaixo da média da carteira variável'
    return 'Carteira equilibrada'


def _build_contribution_action(s, monthly, alloc_map, fii_yield) -> str:
    ro = s.get('reserve_opportunity') or 0
    if ro < RO_TARGET:
        return f'1º R${RO_TARGET - ro:,.2f} → R.O. | Resto → FII de maior yield'
    last_real = next((m['real'] for m in reversed(monthly) if m.get('real')), None)
    if last_real and last_real < META_MENSAL:
        return 'Aportar em SELIC até retomar a META'
    if fii_yield:
        best = max(fii_yield, key=lambda k: fii_yield[k])
        return f'Aportar em {best.upper()} (melhor yield) ou ativo mais subrepresentado'
    return 'Distribuir igualmente entre FIIs'


def _last_month_total_dividends(monthly) -> float:
    last = next((m for m in reversed(monthly) if m.get('total_dividends')), None)
    return last['total_dividends'] if last else 0.0


def _build_dividend_reinvest_message(s, fii_yield, alloc_map) -> str:
    best_fii = max(fii_yield, key=lambda k: fii_yield[k]) if fii_yield else None
    if best_fii:
        return f'Reinvestir em {best_fii.upper()} maximiza o efeito dos juros compostos nos dividendos'
    return 'Considere reinvestir nos FIIs para potencializar os dividendos mensais'


def _build_dividend_reinvest_action(fii_yield, alloc_map) -> str:
    if not fii_yield:
        return 'Reinvestir nos FIIs existentes'
    best = max(fii_yield, key=lambda k: fii_yield[k])
    second = sorted(fii_yield.keys(), key=lambda k: fii_yield[k], reverse=True)
    if len(second) >= 2:
        return f'50% em {best.upper()} + 50% em {second[1].upper()} para diversificar'
    return f'Reinvestir em {best.upper()}'


if __name__ == '__main__':
    app.run(debug=True, port=5000)
