import pandas as pd
import numpy as np
import json

# ============================================================
# READ DATA
# ============================================================
plan_df = pd.read_excel('C:/Users/Administrator/Desktop/山东公司电量计划（包利润使用）.xlsx', header=0)

wind_raw = pd.read_excel('C:/Users/Administrator/Desktop/2026-04_山东公司_风电_生产月报明细.xlsx', header=None, skiprows=3)
wind_raw.columns = ['项目名称','项目编码','设计容量','期末容量','月发电量','月上网电量','月计划','月完成率',
                   '1-本月计划','1-本月完成率','年计划','年完成率','月下网电量','月限电损失电量','月故障损失电量',
                   '月受累损失电量','月均可利用率','年均可利用率','月时长可利用率','年时长可利用率',
                   '月平均风速','年平均风速','累计发电量','累计上网电量','累计下网电量','累计限电损失电量',
                   '累计故障损失电量','累计受累损失电量','上网月小时数','发电月小时数','上网年小时数','发电年小时数','风机型号','台数']

solar_raw = pd.read_excel('C:/Users/Administrator/Desktop/2026-04_山东公司_光伏_生产月报明细.xlsx', header=None, skiprows=3)
solar_raw.columns = ['项目名称','项目编码','设计容量','期末容量','月发电量','月上网电量','月计划','月完成率',
                    '1-本月计划','1-本月完成率','年计划','年完成率','月下网电量','月限电损失电量','月故障损失电量',
                    '月受累损失电量','月累计辐照量','年累计辐照量','累计发电量','累计上网电量',
                    '累计下网电量','累计限电损失电量','累计故障损失电量','累计受累损失电量',
                    '上网月小时数','发电月小时数','上网年小时数','发电年小时数','并网台数']

num_cols_w = ['设计容量','月发电量','月上网电量','月计划','月下网电量','月限电损失电量','月故障损失电量',
              '月受累损失电量','月平均风速','年平均风速','累计发电量','累计上网电量','累计下网电量',
              '累计限电损失电量','累计故障损失电量','累计受累损失电量']
for c in num_cols_w:
    wind_raw[c] = pd.to_numeric(wind_raw[c], errors='coerce').fillna(0)

num_cols_s = ['设计容量','月发电量','月上网电量','月计划','月下网电量','月限电损失电量','月故障损失电量',
              '月受累损失电量','月累计辐照量','年累计辐照量','累计发电量','累计上网电量','累计下网电量',
              '累计限电损失电量','累计故障损失电量','累计受累损失电量']
for c in num_cols_s:
    solar_raw[c] = pd.to_numeric(solar_raw[c], errors='coerce').fillna(0)

def parse_pct(s):
    if pd.isna(s): return 0.0
    s = str(s).replace('%','').strip()
    try: return float(s)/100
    except: return 0.0

for c in ['月完成率','1-本月完成率','年完成率','月均可利用率','年均可利用率','月时长可利用率','年时长可利用率']:
    wind_raw[c] = wind_raw[c].apply(parse_pct)
for c in ['月完成率','1-本月完成率','年完成率']:
    solar_raw[c] = solar_raw[c].apply(parse_pct)

for df_temp in [wind_raw, solar_raw]:
    df_temp['1-本月计划'] = pd.to_numeric(df_temp['1-本月计划'], errors='coerce').fillna(0)
    df_temp['年计划'] = pd.to_numeric(df_temp['年计划'], errors='coerce').fillna(0)

wind_stations = wind_raw[wind_raw['项目名称'] != '山东公司合计'].copy()
wind_total = wind_raw[wind_raw['项目名称'] == '山东公司合计'].iloc[0]
solar_stations = solar_raw[solar_raw['项目名称'] != '山东公司合计'].copy()
solar_total = solar_raw[solar_raw['项目名称'] == '山东公司合计'].iloc[0]

# ============================================================
# PROCESS PLAN DATA
# ============================================================
# FIX: Use the plan values directly from the actual monthly reports (more reliable)
# 月计划, 1-本月计划, 年计划 are already in the actual report files
# For per-station plan matching, merge duplicate project names in plan_lookup
plan_lookup = {}
for idx, row in plan_df.iterrows():
    name = str(row['项目名称']).strip() if pd.notna(row['项目名称']) else ''
    if not name: continue
    yitai = str(row['一级业态']).strip() if pd.notna(row['一级业态']) else ''
    status = str(row['项目状态']).strip() if pd.notna(row['项目状态']) else ''
    entry = {
        'yitai': yitai, 'status': status,
        'capacity': row['容量'] if pd.notna(row['容量']) else 0,
        'plan_4yue': row['4月'] if pd.notna(row['4月']) else 0,
        'plan_cum_4': row['1季度合计'] + (row['4月'] if pd.notna(row['4月']) else 0),
        'plan_year': row['年度合计'] if pd.notna(row['年度合计']) else 0,
    }
    if name in plan_lookup:
        old = plan_lookup[name]
        old['plan_4yue'] += entry['plan_4yue']
        old['plan_cum_4'] += entry['plan_cum_4']
        old['plan_year'] += entry['plan_year']
        old['capacity'] += entry['capacity']
    else:
        plan_lookup[name] = entry

total_plan_row = plan_df.iloc[0]
plan_total = {
    'plan_4yue': total_plan_row['4月'] if pd.notna(total_plan_row['4月']) else 0,
    'plan_cum_4': total_plan_row['1季度合计'] + (total_plan_row['4月'] if pd.notna(total_plan_row['4月']) else 0),
    'plan_year': total_plan_row['年度合计'] if pd.notna(total_plan_row['年度合计']) else 0,
}

wind_plan_items = {k:v for k,v in plan_lookup.items() if v['yitai']=='风电'}
solar_plan_items = {k:v for k,v in plan_lookup.items() if v['yitai']=='光伏'}

wind_plan_total_4yue = sum(v['plan_4yue'] for v in wind_plan_items.values())
wind_plan_total_cum4 = sum(v['plan_cum_4'] for v in wind_plan_items.values())
wind_plan_total_year = sum(v['plan_year'] for v in wind_plan_items.values())

solar_plan_total_4yue = sum(v['plan_4yue'] for v in solar_plan_items.values())
solar_plan_total_cum4 = sum(v['plan_cum_4'] for v in solar_plan_items.values())
solar_plan_total_year = sum(v['plan_year'] for v in solar_plan_items.values())

# ============================================================
# BUILD STATION-LEVEL DATA
# ============================================================
all_stations = []

for _, row in wind_stations.iterrows():
    name = str(row['项目名称']).strip()
    plan = plan_lookup.get(name, {'plan_4yue': 0, 'plan_cum_4': 0, 'plan_year': 0, 'yitai': '风电', 'status': '未知'})
    am = float(row['月上网电量'])
    ac = float(row['累计上网电量'])
    pm = float(plan['plan_4yue'])
    pc = float(plan['plan_cum_4'])
    py = float(plan['plan_year'])
    all_stations.append({
        'name': name, 'yitai': '风电', 'status': plan['status'],
        'capacity': float(row['设计容量']),
        'actual_month': am, 'actual_cum': ac, 'plan_month': pm, 'plan_cum': pc, 'plan_year': py,
        'over_under_month': am - pm, 'over_under_cum': ac - pc,
        'rate_month': am/pm if pm>0 else 0, 'rate_cum': ac/pc if pc>0 else 0, 'rate_year': ac/py if py>0 else 0,
        'gen_month': float(row['月发电量']), 'gen_cum': float(row['累计发电量']),
        'down_month': float(row['月下网电量']), 'down_cum': float(row['累计下网电量']),
        'limit_month': float(row['月限电损失电量']), 'limit_cum': float(row['累计限电损失电量']),
        'wind_month': float(row['月平均风速']), 'wind_year': float(row['年平均风速']),
        'irr_month': 0, 'irr_year': 0,
    })

for _, row in solar_stations.iterrows():
    name = str(row['项目名称']).strip()
    plan = plan_lookup.get(name, {'plan_4yue': 0, 'plan_cum_4': 0, 'plan_year': 0, 'yitai': '光伏', 'status': '未知'})
    am = float(row['月上网电量'])
    ac = float(row['累计上网电量'])
    pm = float(plan['plan_4yue'])
    pc = float(plan['plan_cum_4'])
    py = float(plan['plan_year'])
    all_stations.append({
        'name': name, 'yitai': '光伏', 'status': plan['status'],
        'capacity': float(row['设计容量']),
        'actual_month': am, 'actual_cum': ac, 'plan_month': pm, 'plan_cum': pc, 'plan_year': py,
        'over_under_month': am - pm, 'over_under_cum': ac - pc,
        'rate_month': am/pm if pm>0 else 0, 'rate_cum': ac/pc if pc>0 else 0, 'rate_year': ac/py if py>0 else 0,
        'gen_month': float(row['月发电量']), 'gen_cum': float(row['累计发电量']),
        'down_month': float(row['月下网电量']), 'down_cum': float(row['累计下网电量']),
        'limit_month': float(row['月限电损失电量']), 'limit_cum': float(row['累计限电损失电量']),
        'wind_month': 0, 'wind_year': 0,
        'irr_month': float(row['月累计辐照量']), 'irr_year': float(row['年累计辐照量']),
    })

output = {
    'plan_total': plan_total,
    'wind_plan_total': {'plan_4yue': wind_plan_total_4yue, 'plan_cum_4': wind_plan_total_cum4, 'plan_year': wind_plan_total_year},
    'solar_plan_total': {'plan_4yue': solar_plan_total_4yue, 'plan_cum_4': solar_plan_total_cum4, 'plan_year': solar_plan_total_year},
    'wind_total': {
        'actual_month': float(wind_total['月上网电量']), 'actual_cum': float(wind_total['累计上网电量']),
        'gen_month': float(wind_total['月发电量']), 'gen_cum': float(wind_total['累计发电量']),
        'down_month': float(wind_total['月下网电量']), 'down_cum': float(wind_total['累计下网电量']),
        'limit_month': float(wind_total['月限电损失电量']), 'limit_cum': float(wind_total['累计限电损失电量']),
        'wind_month': float(wind_total['月平均风速']), 'wind_year': float(wind_total['年平均风速']),
    },
    'solar_total': {
        'actual_month': float(solar_total['月上网电量']), 'actual_cum': float(solar_total['累计上网电量']),
        'gen_month': float(solar_total['月发电量']), 'gen_cum': float(solar_total['累计发电量']),
        'down_month': float(solar_total['月下网电量']), 'down_cum': float(solar_total['累计下网电量']),
        'limit_month': float(solar_total['月限电损失电量']), 'limit_cum': float(solar_total['累计限电损失电量']),
        'irr_month': float(solar_total['月累计辐照量']), 'irr_year': float(solar_total['年累计辐照量']),
    },
    'stations': all_stations,
}

with open('analysis_data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print("OK - data saved")
