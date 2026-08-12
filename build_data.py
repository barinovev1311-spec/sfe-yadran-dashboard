#!/usr/bin/env python3
"""
build_data.py — пересборка данных дашборда SFE Ядран из исходных Excel-файлов.

Запуск:
    python3 build_data.py --census ПЕРЕПИСЬ.xlsx --sfe МОДЕЛЬ_SFE.xlsx --outdir data

Оба аргумента опциональны — можно пересобрать только один источник:
    python3 build_data.py --census НОВАЯ_ПЕРЕПИСЬ.xlsx        # обновить только территории
    python3 build_data.py --sfe НОВАЯ_МОДЕЛЬ_SFE.xlsx         # обновить только вкладку МП

Требования: pip install pandas openpyxl
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    import pandas as pd
except ImportError:
    sys.exit("Нужен pandas: pip install pandas openpyxl")

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

TO_MAP = {
    '01. МЕНЕЕ 0,75 МЛН. РУБ.': 0.5, '02. ОТ 0,75 ДО 1,5 МЛН. РУБ.': 1.125,
    '03. ОТ 1,5 ДО 2,25 МЛН. РУБ.': 1.875, '04. ОТ 2,25 ДО 3 МЛН. РУБ.': 2.625,
    '05. ОТ 3 ДО 3,75 МЛН. РУБ.': 3.375, '06. ОТ 3,75 ДО 4,5 МЛН. РУБ.': 4.125,
    '07. ОТ 4,5 ДО 5,25 МЛН. РУБ.': 4.875, '08. ОТ 5,25 ДО 7,5 МЛН. РУБ.': 6.375,
    '09. ОТ 7,5 ДО 10 МЛН. РУБ.': 8.75, '10. БОЛЕЕ 10 МЛН. РУБ.': 12.5,
}
TO_RANK = {k: i + 1 for i, k in enumerate(sorted(TO_MAP.keys()))}
CHECK_MAP = {
    'ОТ 100 ДО 200 РУБ.': 150, 'ОТ 200 ДО 300 РУБ.': 250, 'ОТ 300 ДО 400 РУБ.': 350,
    'ОТ 400 ДО 500 РУБ.': 450, 'ОТ 500 ДО 600 РУБ.': 550, 'БОЛЕЕ 600 РУБ.': 700,
}
TRANS = {
    'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'e','ж':'zh','з':'z','и':'i',
    'й':'y','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t',
    'у':'u','ф':'f','х':'h','ц':'ts','ч':'ch','ш':'sh','щ':'sch','ъ':'','ы':'y','ь':'',
    'э':'e','ю':'yu','я':'ya',' ':'_','-':'_',
}


def translit_slug(s):
    s = s.lower().strip()
    out = []
    for ch in s:
        if ch in TRANS:
            out.append(TRANS[ch])
        elif ch.isalnum():
            out.append(ch)
        else:
            out.append('_')
    slug = ''.join(out)
    while '__' in slug:
        slug = slug.replace('__', '_')
    return slug.strip('_')


def score_census(df):
    """Добавляет колонки-скоринг к датафрейму переписи. Возвращает тот же df."""
    df = df.copy()
    df['to_val'] = df['КАТЕГОРИЯ RNC ПО ВАЛОВОМУ ТО'].map(TO_MAP)
    df['to_rank'] = df['КАТЕГОРИЯ RNC ПО ВАЛОВОМУ ТО'].map(TO_RANK)
    to_label = {k: k.split('. ', 1)[1].replace('РУБ.', 'руб.').replace('МЛН', 'млн') for k in TO_MAP}
    df['to_label'] = df['КАТЕГОРИЯ RNC ПО ВАЛОВОМУ ТО'].map(to_label)
    df['check_val'] = df['ОЦЕНКА СРЕДНЕГО ЧЕКА ТОЧКИ'].map(CHECK_MAP)
    df['is_chain'] = (df['СЕТЬ'] != '-').astype(int)

    df['n_to'] = (df['to_rank'] - 1) / 9 * 100
    df['n_check'] = (df['check_val'] - df['check_val'].min()) / (df['check_val'].max() - df['check_val'].min()) * 100
    lp_pct = df['ЛП, %'] * 100
    df['n_lp'] = (lp_pct - lp_pct.min()) / (lp_pct.max() - lp_pct.min()) * 100
    df['score'] = (df['n_to'] * 0.5 + df['n_check'] * 0.25 + df['n_lp'] * 0.15 + df['is_chain'] * 10).round(1).clip(0, 100)

    def tier(s):
        if s >= 70: return 'A'
        if s >= 50: return 'B'
        if s >= 30: return 'C'
        return 'D'
    df['tier'] = df['score'].apply(tier)
    return df


def process_census(path, outdir):
    print(f'[перепись] читаю {path}')
    df = pd.read_excel(path)
    required = ['ИНН', 'ID RNC', 'СЕТЬ', 'СЕТЬ, ЯДРАН', 'СЕТЬ, ЯДРАН ДЕТАЛИЗАЦИЯ', 'ЮРИДИЧЕСКОЕ ЛИЦО',
                'АДРЕС', 'СУБЪЕКТ ФЕДЕРАЦИИ', 'НАСЕЛЕННЫЙ ПУНКТ', 'КАТЕГОРИЯ RNC ПО ВАЛОВОМУ ТО',
                'ОЦЕНКА СРЕДНЕГО ЧЕКА ТОЧКИ', 'ЛП, %', 'БАД, %', 'ПРОЧИЙ АССОРТИМЕНТ, %']
    missing = [c for c in required if c not in df.columns]
    if missing:
        sys.exit(f'[перепись] В файле не хватает колонок: {missing}\nПроверь, что структура файла не изменилась.')

    df['to_rank'] = df['КАТЕГОРИЯ RNC ПО ВАЛОВОМУ ТО'].map(TO_RANK)
    df['check_val'] = df['ОЦЕНКА СРЕДНЕГО ЧЕКА ТОЧКИ'].map(CHECK_MAP)
    lp_pct_all = df['ЛП, %'] * 100

    os.makedirs(outdir, exist_ok=True)
    index = []
    seen = {}
    # компактный формат: словари (net/grp/org/city) + строки-массивы без ключей.
    # score/tier/toCat/toVal больше НЕ хранятся — считаются в браузере по _meta.json.
    for region, g in df.groupby('СУБЪЕКТ ФЕДЕРАЦИИ'):
        slug = translit_slug(region)
        if slug in seen:
            seen[slug] += 1
            slug = f'{slug}_{seen[slug]}'
        else:
            seen[slug] = 0

        net_dict, grp_dict, org_dict, city_dict = {}, {}, {}, {}
        def idx_of(d, val):
            val = val if val and val != '-' else ''
            if val not in d:
                d[val] = len(d)
            return d[val]

        rows = []
        for _, r in g.iterrows():
            ni = idx_of(net_dict, r['СЕТЬ'])
            gi = idx_of(grp_dict, r['СЕТЬ, ЯДРАН'])
            oi = idx_of(org_dict, r['ЮРИДИЧЕСКОЕ ЛИЦО'])
            ci = idx_of(city_dict, r['НАСЕЛЕННЫЙ ПУНКТ'])
            rows.append([
                int(r['ID RNC']), ni, gi, oi, r['АДРЕС'], ci,
                int(r['to_rank']), int(r['check_val']),
                round(float(r['ЛП, %']) * 100, 1), round(float(r['БАД, %']) * 100, 1),
                round(float(r['ПРОЧИЙ АССОРТИМЕНТ, %']) * 100, 1),
            ])

        payload = {
            'd': {
                'n': sorted(net_dict, key=net_dict.get), 'g': sorted(grp_dict, key=grp_dict.get),
                'o': sorted(org_dict, key=org_dict.get), 'c': sorted(city_dict, key=city_dict.get),
            },
            'r': rows,
        }
        with open(os.path.join(outdir, f'{slug}.json'), 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, separators=(',', ':'))

        scored = score_census(g)
        cities = sorted(g['НАСЕЛЕННЫЙ ПУНКТ'].unique().tolist())
        index.append({'region': region, 'slug': slug, 'count': len(g), 'avgScore': round(scored['score'].mean(), 1),
                       'tierA': int((scored['tier'] == 'A').sum()), 'cities': cities})
    index.sort(key=lambda x: x['region'])
    with open(os.path.join(outdir, '_index.json'), 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, separators=(',', ':'))

    to_label = {k: k.split('. ', 1)[1].replace('РУБ.', 'руб.').replace('МЛН', 'млн') for k in TO_MAP}
    meta = {
        'toLabels': [to_label[k] for k in sorted(TO_MAP, key=TO_RANK.get)],
        'toVals': [TO_MAP[k] for k in sorted(TO_MAP, key=TO_RANK.get)],
        'checkMin': min(CHECK_MAP.values()), 'checkMax': max(CHECK_MAP.values()),
        'lpMin': round(float(lp_pct_all.min()), 3), 'lpMax': round(float(lp_pct_all.max()), 3),
    }
    with open(os.path.join(outdir, '_meta.json'), 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)

    df = score_census(df)  # для линковки с SFE, если понадобится в этом же запуске
    print(f'[перепись] готово: {len(index)} регионов, {len(df)} точек -> {outdir}/ (компактный формат)')
    return df


def read_pivot_cache(xlsx_path):
    """Достаёт данные из скрытого pivot cache файла (первый найденный)."""
    with zipfile.ZipFile(xlsx_path) as z:
        cache_defs = [n for n in z.namelist() if re.match(r'xl/pivotCache/pivotCacheDefinition\d*\.xml', n)]
        cache_recs = [n for n in z.namelist() if re.match(r'xl/pivotCache/pivotCacheRecords\d*\.xml', n)]
        if not cache_defs or not cache_recs:
            return None
        def_xml = z.read(cache_defs[0])
        rec_xml = z.read(cache_recs[0])

    root = ET.fromstring(def_xml)
    field_names, field_shared = [], []
    for f in root.iter(NS + 'cacheField'):
        field_names.append(f.get('name'))
        items_el = f.find(NS + 'sharedItems')
        shared = [i.get('v') for i in items_el] if items_el is not None and len(items_el) > 0 else None
        field_shared.append(shared)

    rroot = ET.fromstring(rec_xml)
    records = []
    for rec in rroot:
        row = []
        for i, cell in enumerate(rec):
            tag = cell.tag.replace(NS, '')
            if tag == 'x':
                idx = cell.get('v')
                val = field_shared[i][int(idx)] if idx is not None and field_shared[i] else None
            elif tag == 'm':
                val = None
            else:
                val = cell.get('v')
            row.append(val)
        records.append(row)
    return pd.DataFrame(records, columns=field_names)


def process_sfe(path, outdir, census_df=None):
    print(f'[SFE] читаю {path}')
    sfe = read_pivot_cache(path)
    if sfe is None:
        sys.exit(
            '[SFE] Не нашёл pivot cache внутри файла. Похоже, структура выгрузки изменилась '
            '(например, файл больше не сводная таблица, а плоский экспорт). '
            'Если это так — пришли пример новой структуры, скрипт нужно будет адаптировать под неё.'
        )
    required = ['МП', 'БРИК', 'Менеджер', 'ID ТОЧКИ_RNC', 'Целевой Q3', 'Целевой в след цикле Q4',
                'Наличие визита Q1', 'Потенциал', 'ЮРИДИЧЕСКОЕ ЛИЦО', 'АДРЕС АПТЕКИ', 'НАСЕЛЕННЫЙ ПУНКТ',
                'СЕТЬ, ЯДРАН']
    missing = [c for c in required if c not in sfe.columns]
    if missing:
        sys.exit(f'[SFE] В pivot cache не хватает полей: {missing}\nПроверь структуру исходной сводной таблицы в Excel.')

    # определяем помесячные колонки продаж динамически (вдруг период изменится)
    db_cols = sorted([c for c in sfe.columns if re.match(r'Продажи ДБ 0\d 25 Val', c)])
    ac_cols = sorted([c for c in sfe.columns if re.match(r'Продажи AC 0\d 25 Val', c)])
    q_cols = ['Продажи ДБ Q1 25 Val', 'Продажи ДБ Q2 25 Val', 'Продажи АС Q1 25 Val', 'Продажи АС Q2 25 Val']
    for c in db_cols + ac_cols + q_cols:
        if c in sfe.columns:
            sfe[c] = pd.to_numeric(sfe[c], errors='coerce').fillna(0)

    sfe['ID ТОЧКИ_RNC'] = pd.to_numeric(sfe['ID ТОЧКИ_RNC'], errors='coerce')
    sfe['sales_q1'] = sfe.get('Продажи ДБ Q1 25 Val', 0) + sfe.get('Продажи АС Q1 25 Val', 0)
    sfe['sales_q2'] = sfe.get('Продажи ДБ Q2 25 Val', 0) + sfe.get('Продажи АС Q2 25 Val', 0)
    sfe['sales_total'] = sfe['sales_q1'] + sfe['sales_q2']
    sfe['visited'] = sfe['Наличие визита Q1'].fillna('').str.lower() == 'да'
    sfe['targetQ3'] = sfe['Целевой Q3'].fillna('').str.lower() == 'да'
    sfe['targetQ4'] = sfe['Целевой в след цикле Q4'].fillna('').str.lower() == 'да'
    sfe['plan_visited'] = sfe['targetQ3'] & sfe['visited']
    sfe['extra_visited'] = (~sfe['targetQ3']) & sfe['visited']
    sfe['missed_visit'] = sfe['targetQ3'] & (~sfe['visited'])

    # линковка со скорингом переписи по точному ID точки
    score_map, tier_map = {}, {}
    if census_df is not None and 'ID RNC' in census_df.columns:
        if 'score' not in census_df.columns:
            census_df = score_census(census_df)
        score_map = census_df.set_index('ID RNC')['score'].to_dict()
        tier_map = census_df.set_index('ID RNC')['tier'].to_dict()
    sfe['census_score'] = sfe['ID ТОЧКИ_RNC'].map(score_map)
    sfe['census_tier'] = sfe['ID ТОЧКИ_RNC'].map(tier_map)
    matched = sfe['census_score'].notna().sum()
    print(f'[SFE] сопоставлено с перепиской по ID точки: {matched}/{len(sfe)}')

    assigned = sfe[sfe['МП'].notna()].copy()
    pot_map = {'1-3(Низкий)': 'Низкий', '2-6(Средний)': 'Средний', '7-10(Высокий)': 'Высокий'}
    assigned['pot'] = assigned['Потенциал'].map(pot_map)
    print(f'[SFE] точек с назначенным МП: {len(assigned)}/{len(sfe)}')

    def monthly_sum(row):
        vals = []
        for i in range(1, 7):
            db_c = f'Продажи ДБ 0{i} 25 Val'
            ac_c = f'Продажи AC 0{i} 25 Val'
            v = float(row.get(db_c, 0) or 0) + float(row.get(ac_c, 0) or 0)
            vals.append(round(v, 1))
        return vals

    pot_code = {'Низкий': 0, 'Средний': 1, 'Высокий': 2}
    mp_dict, org_dict, net_dict, city_dict = {}, {}, {}, {}
    def idx_of(d, val):
        val = val if val else ''
        if val not in d:
            d[val] = len(d)
        return d[val]

    rows = []
    for _, r in assigned.iterrows():
        mi = idx_of(mp_dict, r['МП'])
        oi = idx_of(org_dict, r['ЮРИДИЧЕСКОЕ ЛИЦО'] or '')
        ni = idx_of(net_dict, r['СЕТЬ, ЯДРАН'] if r['СЕТЬ, ЯДРАН'] not in ('-', None) else '')
        ci = idx_of(city_dict, r['НАСЕЛЕННЫЙ ПУНКТ'] or '')
        pot = pot_code.get(r['pot'], -1)
        cs = round(float(r['census_score']), 1) if pd.notna(r['census_score']) else None
        rows.append([mi, oi, r['АДРЕС АПТЕКИ'] or '', ci, ni, pot,
                     1 if r['targetQ3'] else 0, 1 if r['visited'] else 0,
                     round(float(r['sales_q1']), 1), round(float(r['sales_q2']), 1), cs])

    points_payload = {
        'd': {'mp': sorted(mp_dict, key=mp_dict.get), 'o': sorted(org_dict, key=org_dict.get),
              'n': sorted(net_dict, key=net_dict.get), 'c': sorted(city_dict, key=city_dict.get)},
        'r': rows,
    }
    with open(os.path.join(outdir, 'mp_points.json'), 'w', encoding='utf-8') as f:
        json.dump(points_payload, f, ensure_ascii=False, separators=(',', ':'))

    reps = []
    for mp, g in assigned.groupby('МП'):
        targeted = int(g['targetQ3'].sum())
        plan_visited = int(g['plan_visited'].sum())
        reps.append({
            'mp': mp, 'mgr': g['Менеджер'].mode().iat[0] if len(g['Менеджер'].mode()) else '',
            'bricks': sorted(g['БРИК'].dropna().unique().tolist()),
            'points': len(g), 'targetedQ3': targeted, 'targetedQ4': int(g['targetQ4'].sum()),
            'planVisited': plan_visited, 'extraVisited': int(g['extra_visited'].sum()), 'missedVisit': int(g['missed_visit'].sum()),
            'visitRate': round(plan_visited / targeted * 100, 1) if targeted else None,
            'salesQ1': round(g['sales_q1'].sum(), 1), 'salesQ2': round(g['sales_q2'].sum(), 1),
            'salesTotal': round(g['sales_total'].sum(), 1),
            'growthPct': round((g['sales_q2'].sum() - g['sales_q1'].sum()) / g['sales_q1'].sum() * 100, 1) if g['sales_q1'].sum() > 0 else None,
            'avgCensusScore': round(g['census_score'].mean(), 1) if g['census_score'].notna().any() else None,
            'highPotNotTargeted': int(((g['Потенциал'] == '7-10(Высокий)') & (~g['targetQ3'])).sum()),
            'monthly': [round(float(g[f'Продажи ДБ 0{i} 25 Val'].sum() if f'Продажи ДБ 0{i} 25 Val' in g else 0) +
                               float(g[f'Продажи AC 0{i} 25 Val'].sum() if f'Продажи AC 0{i} 25 Val' in g else 0), 1) for i in range(1, 7)],
        })
    reps.sort(key=lambda x: -x['salesTotal'])
    n = len(reps)
    for i, r in enumerate(sorted(reps, key=lambda x: x['salesTotal'])):
        r['salesPct'] = round((i + 1) / n * 100, 1)
    vr = [r for r in reps if r['visitRate'] is not None]
    for i, r in enumerate(sorted(vr, key=lambda x: x['visitRate'])):
        r['visitPct'] = round((i + 1) / len(vr) * 100, 1)
    for r in reps:
        r.setdefault('visitPct', None)
    with open(os.path.join(outdir, 'mp_reps.json'), 'w', encoding='utf-8') as f:
        json.dump(reps, f, ensure_ascii=False, separators=(',', ':'))

    bricks = []
    for brick, g in sfe.groupby('БРИК'):
        bricks.append({
            'brick': brick, 'totalPoints': len(g), 'targetedQ3': int(g['targetQ3'].sum()),
            'reps': sorted(g['МП'].dropna().unique().tolist()), 'salesTotal': round(g['sales_total'].sum(), 1),
            'potHigh': int((g['Потенциал'] == '7-10(Высокий)').sum()),
            'potMed': int((g['Потенциал'] == '2-6(Средний)').sum()),
            'potLow': int((g['Потенциал'] == '1-3(Низкий)').sum()),
            'highPotNotTargeted': int(((g['Потенциал'] == '7-10(Высокий)') & (~g['targetQ3'])).sum()),
        })
    with open(os.path.join(outdir, 'mp_bricks.json'), 'w', encoding='utf-8') as f:
        json.dump(bricks, f, ensure_ascii=False, separators=(',', ':'))

    growth_vals = [r['growthPct'] for r in reps if r['growthPct'] is not None]
    company = {
        'avgVisitRate': round(sum(r['visitRate'] for r in vr) / len(vr), 1) if vr else None,
        'avgSalesPerRep': round(sum(r['salesTotal'] for r in reps) / n, 1) if n else 0,
        'avgGrowthPct': round(sum(growth_vals) / len(growth_vals), 1) if growth_vals else None,
        'totalReps': n, 'totalAssignedPoints': len(assigned), 'totalBricks': len(bricks),
        'totalExtraVisited': int(assigned['extra_visited'].sum()), 'totalMissedVisit': int(assigned['missed_visit'].sum()),
    }
    with open(os.path.join(outdir, 'mp_company.json'), 'w', encoding='utf-8') as f:
        json.dump(company, f, ensure_ascii=False)
    print(f'[SFE] готово: {n} МП, {len(bricks)} бриков -> {outdir}/')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--census', help='путь к файлу переписи аптек (xlsx)')
    ap.add_argument('--sfe', help='путь к файлу модели сегментации SFE (xlsx)')
    ap.add_argument('--outdir', default='data', help='папка для файлов данных (по умолчанию ./data)')
    args = ap.parse_args()

    if not args.census and not args.sfe:
        ap.error('укажи хотя бы --census или --sfe')

    census_df = None
    if args.census:
        census_df = process_census(args.census, args.outdir)
    if args.sfe:
        if census_df is None and os.path.exists(os.path.join(args.outdir, '_index.json')):
            print('[SFE] --census не передан в этом запуске, беру для линковки существующие данные переписи, если понадобятся числа — пересобери --census заново для точных cScore/cTier')
        process_sfe(args.sfe, args.outdir, census_df)

    print('\nГотово. Теперь закоммить папку', args.outdir, 'и запушь в GitHub — Railway передеплоит сайт автоматически.')


if __name__ == '__main__':
    main()
