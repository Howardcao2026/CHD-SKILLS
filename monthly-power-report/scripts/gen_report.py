# -*- coding: utf-8 -*-
import json

with open('analysis_data.json', 'r', encoding='utf-8') as f:
    D = json.load(f)

stations = D['stations']
pt = D['plan_total']
wp = D['wind_plan_total']
sp = D['solar_plan_total']
wt = D['wind_total']
st = D['solar_total']

def fmt(v, dec=2):
    if abs(v) >= 10000:
        return f"{v:,.{dec}f}"
    return f"{v:.{dec}f}"

def pct(v):
    return f"{v*100:.2f}%"

def pct_color(v, threshold_good=1.0):
    if v >= threshold_good:
        return "#16a34a"
    elif v >= 0.9:
        return "#ea580c"
    else:
        return "#dc2626"

def ou_color(v):
    if v > 0: return "#16a34a"
    elif v < 0: return "#dc2626"
    return "#374151"

def ou_arrow(v):
    if v > 0:
        return f'<span style="color:#16a34a">&#9650;+{fmt(v)}</span>'
    elif v < 0:
        return f'<span style="color:#dc2626">&#9660;{fmt(v)}</span>'
    return f'<span style="color:#374151">{fmt(v)}</span>'

# ============================================================
# SECTION 1
# ============================================================
company_actual_month = wt['actual_month'] + st['actual_month']
company_actual_cum = wt['actual_cum'] + st['actual_cum']

def calc_rates(am, ac, pm, pc, py):
    return {
        'ou_m': am - pm, 'ou_c': ac - pc,
        'rate_m': am/pm if pm>0 else 0, 'rate_c': ac/pc if pc>0 else 0, 'rate_y': ac/py if py>0 else 0,
    }

wr = calc_rates(wt['actual_month'], wt['actual_cum'], wp['plan_4yue'], wp['plan_cum_4'], wp['plan_year'])
sr = calc_rates(st['actual_month'], st['actual_cum'], sp['plan_4yue'], sp['plan_cum_4'], sp['plan_year'])
cr = calc_rates(company_actual_month, company_actual_cum, pt['plan_4yue'], pt['plan_cum_4'], pt['plan_year'])

# ============================================================
# SECTION 2: Incremental / Existing (FIXED: cumulative only)
# ============================================================
inc_2025_names = {'招远400MW海光', '莱州土山600MW盐光'}
inc_2026_names = {'刘王庄', '汶上南站', '汶上白石', '日照莒县中北部'}

inc25 = [s for s in stations if s['name'] in inc_2025_names]
inc26 = [s for s in stations if s['name'] in inc_2026_names]
exist = [s for s in stations if s['name'] not in inc_2025_names and s['name'] not in inc_2026_names]

def group_cum(group):
    ac = sum(s['actual_cum'] for s in group)
    pc = sum(s['plan_cum'] for s in group)
    py = sum(s['plan_year'] for s in group)
    return {'actual_cum': ac, 'plan_cum': pc, 'plan_year': py,
            'ou': ac-pc, 'rate': ac/pc if pc>0 else 0, 'rate_y': ac/py if py>0 else 0}

g25 = group_cum(inc25)
g26 = group_cum(inc26)
gex = group_cum(exist)

# ============================================================
# SECTION 3
# ============================================================
top3_over_month = sorted(stations, key=lambda x: -x['over_under_month'])[:3]
top3_under_month = sorted(stations, key=lambda x: x['over_under_month'])[:3]
top3_over_cum = sorted(stations, key=lambda x: -x['over_under_cum'])[:3]
top3_under_cum = sorted(stations, key=lambda x: x['over_under_cum'])[:3]

# ============================================================
# SECTION 4: Curtailment
# ============================================================
def limit_rate(lim, gen):
    return lim / (lim + gen) if (lim + gen) > 0 else 0

wlr_m = limit_rate(wt['limit_month'], wt['gen_month'])
wlr_y = limit_rate(wt['limit_cum'], wt['gen_cum'])
slr_m = limit_rate(st['limit_month'], st['gen_month'])
slr_y = limit_rate(st['limit_cum'], st['gen_cum'])
clr_m = limit_rate(wt['limit_month']+st['limit_month'], wt['gen_month']+st['gen_month'])
clr_y = limit_rate(wt['limit_cum']+st['limit_cum'], wt['gen_cum']+st['gen_cum'])

station_rates = []
for s in stations:
    station_rates.append({**s, 'lr_m': limit_rate(s['limit_month'], s['gen_month']),
                          'lr_y': limit_rate(s['limit_cum'], s['gen_cum'])})

base_names = {'莱州海风', '招远400MW海光', '莱州土山600MW盐光'}
base_s = [s for s in station_rates if s['name'] in base_names]
b_gen_m = sum(s['gen_month'] for s in base_s)
b_lim_m = sum(s['limit_month'] for s in base_s)
b_gen_c = sum(s['gen_cum'] for s in base_s)
b_lim_c = sum(s['limit_cum'] for s in base_s)

# ============================================================
# SECTION 5
# ============================================================
key_stations = {'枣庄山亭', '莱州海风', '招远400MW海光', '莱州土山600MW盐光', '刘王庄'}
key_data = [s for s in stations if s['name'] in key_stations]

# ============================================================
# SECTION 6: Resources
# ============================================================
wind_resource = [s for s in stations if s['yitai'] == '风电' and s['wind_month'] > 0]
solar_resource = [s for s in stations if s['yitai'] == '光伏' and s['irr_month'] > 0]

# ============================================================
# SECTION 7: Self-consumption (FIXED: 4 columns only)
# ============================================================
for s in stations:
    s['su_m'] = s['gen_month'] - s['actual_month'] + s['down_month']
    s['su_c'] = s['gen_cum'] - s['actual_cum'] + s['down_cum']
    s['sur_m'] = s['su_m'] / s['gen_month'] if s['gen_month'] > 0 else 0
    s['sur_c'] = s['su_c'] / s['gen_cum'] if s['gen_cum'] > 0 else 0

def self_calc(gen_m, gen_c, on_m, on_c, dn_m, dn_c):
    sm = gen_m - on_m + dn_m
    sc = gen_c - on_c + dn_c
    return {'m': sm, 'c': sc, 'rm': sm/gen_m if gen_m>0 else 0, 'rc': sc/gen_c if gen_c>0 else 0}

ws = self_calc(wt['gen_month'], wt['gen_cum'], wt['actual_month'], wt['actual_cum'], wt['down_month'], wt['down_cum'])
ss = self_calc(st['gen_month'], st['gen_cum'], st['actual_month'], st['actual_cum'], st['down_month'], st['down_cum'])
cs = self_calc(wt['gen_month']+st['gen_month'], wt['gen_cum']+st['gen_cum'],
               company_actual_month, company_actual_cum,
               wt['down_month']+st['down_month'], wt['down_cum']+st['down_cum'])

# ============================================================
# SECTION 8: Utilization hours (FIXED formula)
# ============================================================
wind_cap = wt.get('capacity', 150.9) if 'capacity' in wt else 150.9
solar_cap = st.get('capacity', 106.7) if 'capacity' in st else 106.7
total_cap = wind_cap + solar_cap

# Use actual total capacity from stations
wind_cap = sum(s['capacity'] for s in stations if s['yitai']=='风电')
solar_cap = sum(s['capacity'] for s in stations if s['yitai']=='光伏')
total_cap = wind_cap + solar_cap

def gen_hours(gen, cap):
    return gen / cap if cap > 0 else 0

wgh_m = gen_hours(wt['gen_month'], wind_cap)
wgh_c = gen_hours(wt['gen_cum'], wind_cap)
sgh_m = gen_hours(st['gen_month'], solar_cap)
sgh_c = gen_hours(st['gen_cum'], solar_cap)
tgh_m = gen_hours(wt['gen_month']+st['gen_month'], total_cap)
tgh_c = gen_hours(wt['gen_cum']+st['gen_cum'], total_cap)

# ============================================================
# HTML GENERATION
# ============================================================
html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>山东公司2026年4月电量生产分析报告</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family: 'Microsoft YaHei','PingFang SC',sans-serif; background:#f0f2f5; color:#1f2937; line-height:1.6; }}
  .report {{ max-width:1400px; margin:0 auto; padding:20px; }}
  .header {{ background:linear-gradient(135deg,#1e40af,#3b82f6); color:#fff; padding:30px 40px; border-radius:12px; margin-bottom:24px; }}
  .header h1 {{ font-size:28px; margin-bottom:8px; }}
  .header p {{ font-size:14px; opacity:0.85; }}
  .section {{ background:#fff; border-radius:10px; padding:24px 30px; margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
  .section h2 {{ font-size:20px; color:#1e40af; border-left:4px solid #3b82f6; padding-left:12px; margin-bottom:16px; }}
  .section h3 {{ font-size:16px; color:#374151; margin:16px 0 10px; }}
  table {{ width:100%; border-collapse:collapse; margin:12px 0; font-size:13px; }}
  th {{ background:#eff6ff; color:#1e40af; padding:10px 8px; text-align:center; font-weight:600; border:1px solid #dbeafe; white-space:nowrap; }}
  td {{ padding:8px; text-align:center; border:1px solid #e5e7eb; }}
  tr:nth-child(even) {{ background:#f9fafb; }}
  tr:hover {{ background:#eff6ff; }}
  .kpi-row {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(240px,1fr)); gap:16px; margin:16px 0; }}
  .kpi-card {{ background:linear-gradient(135deg,#eff6ff,#dbeafe); border-radius:10px; padding:16px 20px; text-align:center; border:1px solid #bfdbfe; }}
  .kpi-card.good {{ background:linear-gradient(135deg,#d1fae5,#a7f3d0); border-color:#6ee7b7; }}
  .kpi-card.bad {{ background:linear-gradient(135deg,#fee2e2,#fecaca); border-color:#fca5a5; }}
  .kpi-card.warn {{ background:linear-gradient(135deg,#fef3c7,#fde68a); border-color:#fcd34d; }}
  .kpi-label {{ font-size:12px; color:#6b7280; margin-bottom:4px; }}
  .kpi-value {{ font-size:22px; font-weight:700; color:#1e3a5f; }}
  .kpi-sub {{ font-size:12px; color:#6b7280; margin-top:4px; }}
  .chart-box {{ width:100%; height:420px; margin:16px 0; border:1px solid #e5e7eb; border-radius:8px; }}
  .chart-box-sm {{ width:100%; height:350px; margin:12px 0; border:1px solid #e5e7eb; border-radius:8px; }}
  .note {{ background:#fefce8; border-left:4px solid #eab308; padding:12px 16px; margin:12px 0; border-radius:0 8px 8px 0; font-size:13px; color:#92400e; }}
  .green {{ color:#16a34a; }}
  .red {{ color:#dc2626; }}
  .tag {{ display:inline-block; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }}
  .tag-wind {{ background:#dbeafe; color:#1e40af; }}
  .tag-solar {{ background:#fef3c7; color:#92400e; }}
  .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
  @media(max-width:900px) {{ .two-col {{ grid-template-columns:1fr; }} }}
  .footer {{ text-align:center; padding:20px; color:#9ca3af; font-size:12px; }}
</style>
</head>
<body>
<div class="report">

<div class="header">
  <h1>山东公司 2026年4月 电量生产分析报告</h1>
  <p>报告期：2026年4月 | 数据来源：生产月报明细 + 年度电量计划</p>
</div>

<!-- ==================== SECTION 1 ==================== -->
<div class="section">
<h2>一、风电/光伏/公司整体 上网电量与计划对比</h2>

<div class="kpi-row">
  <div class="kpi-card {'good' if cr['rate_m']>=1 else 'bad' if cr['rate_m']<0.9 else 'warn'}">
    <div class="kpi-label">公司整体 月上网电量</div>
    <div class="kpi-value">{fmt(company_actual_month)} 万kWh</div>
    <div class="kpi-sub">计划 {fmt(pt['plan_4yue'])} | {ou_arrow(cr['ou_m'])} | 完成率 <span class="{'green' if cr['rate_m']>=1 else 'red'}">{pct(cr['rate_m'])}</span></div>
  </div>
  <div class="kpi-card {'good' if cr['rate_c']>=1 else 'bad' if cr['rate_c']<0.9 else 'warn'}">
    <div class="kpi-label">公司整体 1-4月累计上网电量</div>
    <div class="kpi-value">{fmt(company_actual_cum)} 万kWh</div>
    <div class="kpi-sub">计划 {fmt(pt['plan_cum_4'])} | {ou_arrow(cr['ou_c'])} | 完成率 <span class="{'green' if cr['rate_c']>=1 else 'red'}">{pct(cr['rate_c'])}</span></div>
  </div>
  <div class="kpi-card {'good' if cr['rate_y']>=0.25 else 'warn'}">
    <div class="kpi-label">公司整体 年度计划完成进度</div>
    <div class="kpi-value">{pct(cr['rate_y'])}</div>
    <div class="kpi-sub">年计划 {fmt(pt['plan_year'])} 万kWh | 时间进度 33.33%</div>
  </div>
</div>

<h3>1.1 风电</h3>
<table>
  <tr><th>指标</th><th>实际上网电量</th><th>计划值</th><th>超/欠发</th><th>完成率</th></tr>
  <tr><td>4月当月</td><td>{fmt(wt['actual_month'])}</td><td>{fmt(wp['plan_4yue'])}</td><td>{ou_arrow(wr['ou_m'])}</td><td style="color:{pct_color(wr['rate_m'])};font-weight:700">{pct(wr['rate_m'])}</td></tr>
  <tr><td>1-4月累计</td><td>{fmt(wt['actual_cum'])}</td><td>{fmt(wp['plan_cum_4'])}</td><td>{ou_arrow(wr['ou_c'])}</td><td style="color:{pct_color(wr['rate_c'])};font-weight:700">{pct(wr['rate_c'])}</td></tr>
  <tr><td>年度计划完成率</td><td colspan="3">年计划 {fmt(wp['plan_year'])} 万kWh</td><td style="color:{pct_color(wr['rate_y'])};font-weight:700">{pct(wr['rate_y'])}</td></tr>
</table>

<h3>1.2 光伏</h3>
<table>
  <tr><th>指标</th><th>实际上网电量</th><th>计划值</th><th>超/欠发</th><th>完成率</th></tr>
  <tr><td>4月当月</td><td>{fmt(st['actual_month'])}</td><td>{fmt(sp['plan_4yue'])}</td><td>{ou_arrow(sr['ou_m'])}</td><td style="color:{pct_color(sr['rate_m'])};font-weight:700">{pct(sr['rate_m'])}</td></tr>
  <tr><td>1-4月累计</td><td>{fmt(st['actual_cum'])}</td><td>{fmt(sp['plan_cum_4'])}</td><td>{ou_arrow(sr['ou_c'])}</td><td style="color:{pct_color(sr['rate_c'])};font-weight:700">{pct(sr['rate_c'])}</td></tr>
  <tr><td>年度计划完成率</td><td colspan="3">年计划 {fmt(sp['plan_year'])} 万kWh</td><td style="color:{pct_color(sr['rate_y'])};font-weight:700">{pct(sr['rate_y'])}</td></tr>
</table>

<div class="note">说明：时间进度为33.33%（4个月/12个月）。</div>
<div id="chart1" class="chart-box"></div>
</div>

<!-- ==================== SECTION 2 (FIXED: cumulative only) ==================== -->
<div class="section">
<h2>二、增量项目 vs 存量项目 累计对比（1-4月）</h2>
<table>
  <tr><th>分类</th><th>包含项目</th><th>累计实际上网(万kWh)</th><th>累计计划(万kWh)</th><th>累计超欠发</th><th>累计完成率</th></tr>
  <tr>
    <td><b>2025年增量</b></td>
    <td>招远海光、莱州盐光</td>
    <td><b>{fmt(g25['actual_cum'])}</b></td>
    <td>{fmt(g25['plan_cum'])}</td>
    <td>{ou_arrow(g25['ou'])}</td>
    <td style="color:{pct_color(g25['rate'])};font-weight:700">{pct(g25['rate'])}</td>
  </tr>
  <tr>
    <td><b>2026年增量</b></td>
    <td>刘王庄、济宁白石、汶上南站、莒县北部</td>
    <td><b>{fmt(g26['actual_cum'])}</b></td>
    <td>{fmt(g26['plan_cum'])}</td>
    <td>{ou_arrow(g26['ou'])}</td>
    <td style="font-weight:700">{pct(g26['rate'])}</td>
  </tr>
  <tr style="font-weight:700;background:#eff6ff">
    <td>存量项目</td>
    <td>其余所有场站</td>
    <td>{fmt(gex['actual_cum'])}</td>
    <td>{fmt(gex['plan_cum'])}</td>
    <td>{ou_arrow(gex['ou'])}</td>
    <td style="color:{pct_color(gex['rate'])}">{pct(gex['rate'])}</td>
  </tr>
</table>
<div class="note">
  2026年增量项目中，刘王庄为"大代小"改造项目，济宁白石=汶上白石，济宁汶上=汶上南站，莒县北部=日照莒县中北部。刘王庄和莒县北部4月尚未有实际发电数据。
</div>
<div id="chart2" class="chart-box-sm"></div>
</div>

<!-- ==================== SECTION 3 ==================== -->
<div class="section">
<h2>三、各场站超欠发情况明细</h2>

<h3>3.1 月度超发/欠发排名</h3>
<div class="two-col">
  <div>
    <h4 style="color:#16a34a">月度超发 TOP3</h4>
    <table>
      <tr><th>排名</th><th>场站</th><th>类型</th><th>实际上网</th><th>计划</th><th>超发量</th><th>完成率</th></tr>
      {''.join(f'<tr><td>{i+1}</td><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{fmt(s["actual_month"])}</td><td>{fmt(s["plan_month"])}</td><td class="green">{fmt(s["over_under_month"])}</td><td class="green">{pct(s["rate_month"])}</td></tr>' for i,s in enumerate(top3_over_month))}
    </table>
  </div>
  <div>
    <h4 style="color:#dc2626">月度欠发 TOP3</h4>
    <table>
      <tr><th>排名</th><th>场站</th><th>类型</th><th>实际上网</th><th>计划</th><th>欠发量</th><th>完成率</th></tr>
      {''.join(f'<tr><td>{i+1}</td><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{fmt(s["actual_month"])}</td><td>{fmt(s["plan_month"])}</td><td class="red">{fmt(s["over_under_month"])}</td><td class="red">{pct(s["rate_month"])}</td></tr>' for i,s in enumerate(top3_under_month))}
    </table>
  </div>
</div>

<h3>3.2 累计超发/欠发排名</h3>
<div class="two-col">
  <div>
    <h4 style="color:#16a34a">累计超发 TOP3</h4>
    <table>
      <tr><th>排名</th><th>场站</th><th>类型</th><th>累计实际上网</th><th>累计计划</th><th>超发量</th><th>完成率</th></tr>
      {''.join(f'<tr><td>{i+1}</td><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{fmt(s["actual_cum"])}</td><td>{fmt(s["plan_cum"])}</td><td class="green">{fmt(s["over_under_cum"])}</td><td class="green">{pct(s["rate_cum"])}</td></tr>' for i,s in enumerate(top3_over_cum))}
    </table>
  </div>
  <div>
    <h4 style="color:#dc2626">累计欠发 TOP3</h4>
    <table>
      <tr><th>排名</th><th>场站</th><th>类型</th><th>累计实际上网</th><th>累计计划</th><th>欠发量</th><th>完成率</th></tr>
      {''.join(f'<tr><td>{i+1}</td><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{fmt(s["actual_cum"])}</td><td>{fmt(s["plan_cum"])}</td><td class="red">{fmt(s["over_under_cum"])}</td><td class="red">{pct(s["rate_cum"])}</td></tr>' for i,s in enumerate(top3_under_cum))}
    </table>
  </div>
</div>

<h3>3.3 全部场站月度超欠发柱形图</h3>
<div id="chart3a" class="chart-box"></div>
<h3>3.4 全部场站累计超欠发柱形图</h3>
<div id="chart3b" class="chart-box"></div>

<h3>3.5 全部场站超欠发明细表</h3>
<div style="overflow-x:auto">
<table style="font-size:12px">
  <tr>
    <th>场站名称</th><th>类型</th><th>容量(MW)</th>
    <th>4月实际上网</th><th>4月计划</th><th>月超欠发</th><th>月完成率</th>
    <th>累计实际上网</th><th>累计计划</th><th>累计超欠发</th><th>累计完成率</th><th>年计划完成率</th>
  </tr>
  {''.join(f'<tr><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{s["capacity"]}</td><td>{fmt(s["actual_month"])}</td><td>{fmt(s["plan_month"])}</td><td style="color:{ou_color(s["over_under_month"])};font-weight:600">{fmt(s["over_under_month"])}</td><td style="color:{pct_color(s["rate_month"])};font-weight:600">{pct(s["rate_month"])}</td><td>{fmt(s["actual_cum"])}</td><td>{fmt(s["plan_cum"])}</td><td style="color:{ou_color(s["over_under_cum"])};font-weight:600">{fmt(s["over_under_cum"])}</td><td style="color:{pct_color(s["rate_cum"])};font-weight:600">{pct(s["rate_cum"])}</td><td style="font-weight:600">{pct(s["rate_year"])}</td></tr>' for s in sorted(stations, key=lambda x: x['yitai']))}
</table>
</div>
</div>

<!-- ==================== SECTION 4 ==================== -->
<div class="section">
<h2>四、限电率分析</h2>
<div class="note">限电率 = 限电量 / (限电量 + 发电量)</div>

<h3>4.1 月度及年度限电率汇总</h3>
<table>
  <tr><th>类别</th><th>月限电率</th><th>月限电量(万kWh)</th><th>月发电量(万kWh)</th><th>年限电率</th><th>年累计限电量</th><th>年累计发电量</th></tr>
  <tr><td><span class="tag tag-wind">风电</span></td>
    <td style="font-weight:700;color:{'red' if wlr_m>0.1 else 'orange' if wlr_m>0.05 else 'green'}">{pct(wlr_m)}</td>
    <td>{fmt(wt['limit_month'])}</td><td>{fmt(wt['gen_month'])}</td>
    <td style="font-weight:700;color:{'red' if wlr_y>0.1 else 'orange' if wlr_y>0.05 else 'green'}">{pct(wlr_y)}</td>
    <td>{fmt(wt['limit_cum'])}</td><td>{fmt(wt['gen_cum'])}</td></tr>
  <tr><td><span class="tag tag-solar">光伏</span></td>
    <td style="font-weight:700;color:{'red' if slr_m>0.1 else 'orange' if slr_m>0.05 else 'green'}">{pct(slr_m)}</td>
    <td>{fmt(st['limit_month'])}</td><td>{fmt(st['gen_month'])}</td>
    <td style="font-weight:700;color:{'red' if slr_y>0.1 else 'orange' if slr_y>0.05 else 'green'}">{pct(slr_y)}</td>
    <td>{fmt(st['limit_cum'])}</td><td>{fmt(st['gen_cum'])}</td></tr>
  <tr style="font-weight:700;background:#eff6ff"><td>山东公司整体</td>
    <td style="color:{'red' if clr_m>0.1 else 'orange' if clr_m>0.05 else 'green'}">{pct(clr_m)}</td>
    <td>{fmt(wt['limit_month']+st['limit_month'])}</td><td>{fmt(wt['gen_month']+st['gen_month'])}</td>
    <td style="color:{'red' if clr_y>0.1 else 'orange' if clr_y>0.05 else 'green'}">{pct(clr_y)}</td>
    <td>{fmt(wt['limit_cum']+st['limit_cum'])}</td><td>{fmt(wt['gen_cum']+st['gen_cum'])}</td></tr>
</table>

<div class="note">环比、同比及上月数据需提供3月限电率和2025年4月限电率数据。</div>

<h3>4.2 基地公司限电率（莱州海风、招远海光、莱州土山盐光）</h3>
<table>
  <tr><th>项目</th><th>月限电率</th><th>年限电率</th><th>月发电量</th><th>月限电量</th></tr>
  {''.join(f'<tr><td>{s["name"]}</td><td style="font-weight:700">{pct(s["lr_m"])}</td><td style="font-weight:700">{pct(s["lr_y"])}</td><td>{fmt(s["gen_month"])}</td><td>{fmt(s["limit_month"])}</td></tr>' for s in base_s)}
  <tr style="font-weight:700;background:#eff6ff"><td>基地公司整体</td>
    <td>{pct(limit_rate(b_lim_m, b_gen_m))}</td><td>{pct(limit_rate(b_lim_c, b_gen_c))}</td>
    <td>{fmt(b_gen_m)}</td><td>{fmt(b_lim_m)}</td></tr>
</table>

<h3>4.3 各场站限电率对比</h3>
<div id="chart4" class="chart-box"></div>
</div>

<!-- ==================== SECTION 5 ==================== -->
<div class="section">
<h2>五、关键大场站累计超欠发情况</h2>
<table>
  <tr><th>场站</th><th>类型</th><th>容量(MW)</th><th>4月实际上网</th><th>4月计划</th><th>月超欠发</th><th>月完成率</th><th>累计实际上网</th><th>累计计划</th><th>累计超欠发</th><th>累计完成率</th><th>年计划完成率</th></tr>
  {''.join(f'<tr><td class="highlight">{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{s["capacity"]}</td><td>{fmt(s["actual_month"])}</td><td>{fmt(s["plan_month"])}</td><td style="color:{ou_color(s["over_under_month"])};font-weight:700">{fmt(s["over_under_month"])}</td><td style="color:{pct_color(s["rate_month"])}">{pct(s["rate_month"])}</td><td>{fmt(s["actual_cum"])}</td><td>{fmt(s["plan_cum"])}</td><td style="color:{ou_color(s["over_under_cum"])};font-weight:700">{fmt(s["over_under_cum"])}</td><td style="color:{pct_color(s["rate_cum"])}">{pct(s["rate_cum"])}</td><td style="font-weight:700">{pct(s["rate_year"])}</td></tr>' for s in key_data)}
</table>
<div id="chart5" class="chart-box-sm"></div>
</div>

<!-- ==================== SECTION 6 ==================== -->
<div class="section">
<h2>六、风光资源情况</h2>
<div class="two-col">
  <div>
    <h3>6.1 风速数据（风电场）</h3>
    <table>
      <tr><th>场站</th><th>4月平均风速(m/s)</th><th>年平均风速(m/s)</th></tr>
      {''.join(f'<tr><td>{s["name"]}</td><td>{s["wind_month"]:.2f}</td><td>{s["wind_year"]:.2f}</td></tr>' for s in wind_resource)}
      <tr style="font-weight:700;background:#eff6ff"><td>风电平均</td><td>{wt["wind_month"]:.2f}</td><td>{wt["wind_year"]:.2f}</td></tr>
    </table>
    <div class="note">同比去年数据需提供2025年同期风速。</div>
  </div>
  <div>
    <h3>6.2 辐照量数据（光伏站）</h3>
    <table>
      <tr><th>场站</th><th>4月累计辐照(h)</th><th>年累计辐照(h)</th></tr>
      {''.join(f'<tr><td>{s["name"]}</td><td>{s["irr_month"]:.2f}</td><td>{s["irr_year"]:.2f}</td></tr>' for s in solar_resource)}
      <tr style="font-weight:700;background:#eff6ff"><td>光伏平均</td><td>{st["irr_month"]:.2f}</td><td>{st["irr_year"]:.2f}</td></tr>
    </table>
    <div class="note">同比去年数据需提供2025年同期辐照量。</div>
  </div>
</div>
<div id="chart6a" class="chart-box-sm"></div>
<div id="chart6b" class="chart-box-sm"></div>
</div>

<!-- ==================== SECTION 7 (FIXED: 4 columns only) ==================== -->
<div class="section">
<h2>七、场用电量与场用电率</h2>
<div class="note">场用电量 = 发电量 - 上网电量 + 下网电量；场用电率 = 场用电量 / 发电量</div>

<h3>7.1 公司整体/风电/光伏 汇总</h3>
<table>
  <tr><th>类别</th><th>月场用电量(万kWh)</th><th>月场用电率</th><th>年累计场用电量</th><th>年累计场用电率</th></tr>
  <tr style="font-weight:700;background:#eff6ff"><td>山东公司整体</td><td>{fmt(cs['m'])}</td><td>{pct(cs['rm'])}</td><td>{fmt(cs['c'])}</td><td>{pct(cs['rc'])}</td></tr>
  <tr><td><span class="tag tag-wind">风电</span></td><td>{fmt(ws['m'])}</td><td>{pct(ws['rm'])}</td><td>{fmt(ws['c'])}</td><td>{pct(ws['rc'])}</td></tr>
  <tr><td><span class="tag tag-solar">光伏</span></td><td>{fmt(ss['m'])}</td><td>{pct(ss['rm'])}</td><td>{fmt(ss['c'])}</td><td>{pct(ss['rc'])}</td></tr>
</table>

<h3>7.2 各场站场用电明细</h3>
<table>
  <tr><th>场站</th><th>类型</th><th>月场用电量(万kWh)</th><th>月场用电率</th><th>年累计场用电量(万kWh)</th><th>年累计场用电率</th></tr>
  {''.join(f'<tr><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{fmt(s["su_m"])}</td><td>{pct(s["sur_m"])}</td><td>{fmt(s["su_c"])}</td><td>{pct(s["sur_c"])}</td></tr>' for s in sorted(stations, key=lambda x: x['yitai']))}
</table>
<div id="chart7" class="chart-box-sm"></div>
</div>

<!-- ==================== SECTION 8 (FIXED formula) ==================== -->
<div class="section">
<h2>八、补充分析：发电利用小时数对标</h2>
<div class="note">发电小时数 = 发电量(万kWh) / 装机容量(万kW)。与外部集团或中电联做小时数对标时使用发电小时数。</div>

<h3>8.1 分类型发电小时数</h3>
<table>
  <tr><th>类别</th><th>装机容量(万kW)</th><th>4月发电量(万kWh)</th><th>月发电小时数(h)</th><th>累计发电量(万kWh)</th><th>累计发电小时数(h)</th></tr>
  <tr><td><span class="tag tag-wind">风电</span></td><td>{wind_cap:.1f}</td><td>{fmt(wt['gen_month'])}</td><td style="font-weight:700">{wgh_m:.1f}</td><td>{fmt(wt['gen_cum'])}</td><td style="font-weight:700">{wgh_c:.1f}</td></tr>
  <tr><td><span class="tag tag-solar">光伏</span></td><td>{solar_cap:.1f}</td><td>{fmt(st['gen_month'])}</td><td style="font-weight:700">{sgh_m:.1f}</td><td>{fmt(st['gen_cum'])}</td><td style="font-weight:700">{sgh_c:.1f}</td></tr>
  <tr style="font-weight:700;background:#eff6ff"><td>公司整体</td><td>{total_cap:.1f}</td><td>{fmt(wt['gen_month']+st['gen_month'])}</td><td>{tgh_m:.1f}</td><td>{fmt(wt['gen_cum']+st['gen_cum'])}</td><td>{tgh_c:.1f}</td></tr>
</table>

<h3>8.2 各场站发电小时数排名（按累计小时数降序）</h3>
<div style="overflow-x:auto">
<table style="font-size:12px">
  <tr><th>场站</th><th>类型</th><th>容量(万kW)</th><th>月发电量</th><th>月发电小时数(h)</th><th>累计发电量</th><th>累计发电小时数(h)</th></tr>
  {''.join(f'<tr><td>{s["name"]}</td><td><span class="tag tag-{"solar" if s["yitai"]=="光伏" else "wind"}">{s["yitai"]}</span></td><td>{s["capacity"]}</td><td>{fmt(s["gen_month"])}</td><td>{gen_hours(s["gen_month"],s["capacity"]):.1f}</td><td>{fmt(s["gen_cum"])}</td><td>{gen_hours(s["gen_cum"],s["capacity"]):.1f}</td></tr>' for s in sorted(stations, key=lambda x: -gen_hours(x['gen_cum'],x['capacity']) if x['capacity']>0 else 0))}
</table>
</div>
<div id="chart8" class="chart-box-sm"></div>
</div>

<div class="footer">
  <p>山东公司 2026年4月 电量生产分析报告 | 生成于 2026-05-08</p>
  <p>数据来源：生产月报明细（风电+光伏）+ 年度电量计划表</p>
</div>

</div>

<script>
// ==================== CHART 1 (FIXED: show labels) ====================
var c1 = echarts.init(document.getElementById('chart1'));
c1.setOption({{
  title: {{ text: '风电/光伏/公司整体 月上网电量 vs 计划（万kWh）', left: 'center', textStyle: {{fontSize:15}} }},
  tooltip: {{ trigger: 'axis', formatter: function(p) {{
    var s = p[0].axisValue;
    var r = s + '<br>';
    p.forEach(function(x) {{ r += x.marker + x.seriesName + ': ' + x.value.toLocaleString() + ' 万kWh<br>'; }});
    if (p.length >= 2) {{
      var diff = p[0].value - p[1].value;
      var rate = (p[0].value / p[1].value * 100).toFixed(1);
      r += '完成率: ' + rate + '%（' + (diff>=0?'+':'') + diff.toLocaleString() + '）';
    }}
    return r;
  }} }},
  legend: {{ data: ['实际上网电量','计划电量'], bottom: 0 }},
  grid: {{ top: 50, bottom: 40, left: 80, right: 30 }},
  xAxis: {{ type: 'category', data: ['风电','光伏','公司整体'] }},
  yAxis: {{ type: 'value', name: '万kWh' }},
  series: [
    {{ name: '实际上网电量', type: 'bar', barWidth: 35, data: [{wt['actual_month']:.1f}, {st['actual_month']:.1f}, {company_actual_month:.1f}],
       itemStyle: {{ color: '#3b82f6' }},
       label: {{ show: true, position: 'top', fontSize: 11, formatter: function(p){{ return p.value.toLocaleString(); }} }} }},
    {{ name: '计划电量', type: 'bar', barWidth: 35, data: [{wp['plan_4yue']:.1f}, {sp['plan_4yue']:.1f}, {pt['plan_4yue']:.1f}],
       itemStyle: {{ color: '#fbbf24' }},
       label: {{ show: true, position: 'top', fontSize: 11, formatter: function(p){{ return p.value.toLocaleString(); }} }} }}
  ]
}});

// ==================== CHART 2 (FIXED: cumulative only) ====================
var c2 = echarts.init(document.getElementById('chart2'));
c2.setOption({{
  title: {{ text: '增量 vs 存量项目 1-4月累计完成率', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  grid: {{ top: 50, bottom: 30, left: 60, right: 40 }},
  xAxis: {{ type: 'category', data: ['2025增量','2026增量','存量项目'] }},
  yAxis: {{ type: 'value', name: '%', max: 120 }},
  series: [{{
    type: 'bar', barWidth: 50,
    data: [
      {{ value: {g25['rate']*100:.1f}, itemStyle: {{ color: '{'#10b981' if g25['rate']>=1 else '#ef4444'}' }} }},
      {{ value: {g26['rate']*100:.1f}, itemStyle: {{ color: '{'#10b981' if g26['rate']>=1 else '#f59e0b'}' }} }},
      {{ value: {gex['rate']*100:.1f}, itemStyle: {{ color: '{'#10b981' if gex['rate']>=1 else '#ef4444'}' }} }}
    ],
    label: {{ show: true, position: 'top', formatter: '{{c}}%', fontSize: 13, fontWeight: 'bold' }}
  }}]
}});

// ==================== CHART 3a ====================
var c3a = echarts.init(document.getElementById('chart3a'));
var names3 = {json.dumps([s['name'] for s in sorted(stations, key=lambda x: x['over_under_month'])], ensure_ascii=False)};
var ou3 = {json.dumps([round(s['over_under_month'],1) for s in sorted(stations, key=lambda x: x['over_under_month'])])};
c3a.setOption({{
  title: {{ text: '各场站4月超欠发电量（万kWh）', left: 'center', textStyle: {{fontSize:15}} }},
  tooltip: {{ trigger: 'axis' }},
  grid: {{ top: 50, bottom: 80, left: 80, right: 30 }},
  xAxis: {{ type: 'category', data: names3, axisLabel: {{ rotate: 55, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: '万kWh' }},
  series: [{{
    type: 'bar', data: ou3.map(function(v){{ return {{ value: v, itemStyle: {{ color: v >= 0 ? '#10b981' : '#ef4444' }} }}; }}),
    label: {{ show: true, position: 'outside', fontSize: 9, formatter: function(p){{ return p.value > 0 ? '+'+p.value : ''+p.value; }} }}
  }}]
}});

// ==================== CHART 3b ====================
var c3b = echarts.init(document.getElementById('chart3b'));
var names3b = {json.dumps([s['name'] for s in sorted(stations, key=lambda x: x['over_under_cum'])], ensure_ascii=False)};
var ou3b = {json.dumps([round(s['over_under_cum'],1) for s in sorted(stations, key=lambda x: x['over_under_cum'])])};
c3b.setOption({{
  title: {{ text: '各场站1-4月累计超欠发电量（万kWh）', left: 'center', textStyle: {{fontSize:15}} }},
  tooltip: {{ trigger: 'axis' }},
  grid: {{ top: 50, bottom: 80, left: 80, right: 30 }},
  xAxis: {{ type: 'category', data: names3b, axisLabel: {{ rotate: 55, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: '万kWh' }},
  series: [{{
    type: 'bar', data: ou3b.map(function(v){{ return {{ value: v, itemStyle: {{ color: v >= 0 ? '#10b981' : '#ef4444' }} }}; }}),
    label: {{ show: true, position: 'outside', fontSize: 9, formatter: function(p){{ return p.value > 0 ? '+'+p.value : ''+p.value; }} }}
  }}]
}});

// ==================== CHART 4 ====================
var c4 = echarts.init(document.getElementById('chart4'));
var sr_names = {json.dumps([s['name'] for s in station_rates if s['lr_m'] > 0.001 or s['lr_y'] > 0.001], ensure_ascii=False)};
var sr_m = {json.dumps([round(s['lr_m']*100,2) for s in station_rates if s['lr_m'] > 0.001 or s['lr_y'] > 0.001])};
var sr_y = {json.dumps([round(s['lr_y']*100,2) for s in station_rates if s['lr_m'] > 0.001 or s['lr_y'] > 0.001])};
c4.setOption({{
  title: {{ text: '各场站限电率对比（%）', left: 'center', textStyle: {{fontSize:15}} }},
  tooltip: {{ trigger: 'axis' }},
  legend: {{ data: ['月限电率','年限电率'], bottom: 0 }},
  grid: {{ top: 50, bottom: 40, left: 50, right: 30 }},
  xAxis: {{ type: 'category', data: sr_names, axisLabel: {{ rotate: 45, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: '%', max: 30 }},
  series: [
    {{ name: '月限电率', type: 'bar', data: sr_m, itemStyle: {{ color: '#ef4444' }},
       label: {{ show: true, position: 'top', fontSize: 9, formatter: '{{c}}%' }} }},
    {{ name: '年限电率', type: 'bar', data: sr_y, itemStyle: {{ color: '#f59e0b' }},
       label: {{ show: true, position: 'top', fontSize: 9, formatter: '{{c}}%' }} }}
  ]
}});

// ==================== CHART 5 ====================
var c5 = echarts.init(document.getElementById('chart5'));
var kn = {json.dumps([s['name'] for s in key_data], ensure_ascii=False)};
c5.setOption({{
  title: {{ text: '关键场站累计超欠发（万kWh）', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  legend: {{ data: ['实际上网','计划'], bottom: 0 }},
  grid: {{ top: 40, bottom: 40, left: 80, right: 30 }},
  xAxis: {{ type: 'category', data: kn, axisLabel: {{ rotate: 20 }} }},
  yAxis: {{ type: 'value', name: '万kWh' }},
  series: [
    {{ name: '实际上网', type: 'bar', data: {json.dumps([round(s['actual_cum'],1) for s in key_data])}, itemStyle: {{color:'#3b82f6'}},
       label: {{ show: true, position: 'top', fontSize: 10 }} }},
    {{ name: '计划', type: 'bar', data: {json.dumps([round(s['plan_cum'],1) for s in key_data])}, itemStyle: {{color:'#fbbf24'}},
       label: {{ show: true, position: 'top', fontSize: 10 }} }}
  ]
}});

// ==================== CHART 6a ====================
var c6a = echarts.init(document.getElementById('chart6a'));
c6a.setOption({{
  title: {{ text: '风电场月平均风速（m/s）', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  legend: {{ data: ['4月风速','年平均风速'], bottom: 0 }},
  grid: {{ top: 40, bottom: 40, left: 50, right: 30 }},
  xAxis: {{ type: 'category', data: {json.dumps([s['name'] for s in wind_resource], ensure_ascii=False)}, axisLabel: {{ rotate: 45, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: 'm/s' }},
  series: [
    {{ name: '4月风速', type: 'bar', data: {json.dumps([round(s['wind_month'],2) for s in wind_resource])}, itemStyle: {{color:'#06b6d4'}} }},
    {{ name: '年平均风速', type: 'line', data: {json.dumps([round(s['wind_year'],2) for s in wind_resource])}, itemStyle: {{color:'#f97316'}}, lineStyle: {{width:2}} }}
  ]
}});

// ==================== CHART 6b ====================
var c6b = echarts.init(document.getElementById('chart6b'));
c6b.setOption({{
  title: {{ text: '光伏场月累计辐照量（h）', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  legend: {{ data: ['4月辐照','年累计辐照'], bottom: 0 }},
  grid: {{ top: 40, bottom: 40, left: 50, right: 30 }},
  xAxis: {{ type: 'category', data: {json.dumps([s['name'] for s in solar_resource], ensure_ascii=False)}, axisLabel: {{ rotate: 20 }} }},
  yAxis: {{ type: 'value', name: 'h' }},
  series: [
    {{ name: '4月辐照', type: 'bar', data: {json.dumps([round(s['irr_month'],2) for s in solar_resource])}, itemStyle: {{color:'#f59e0b'}} }},
    {{ name: '年累计辐照', type: 'line', data: {json.dumps([round(s['irr_year'],2) for s in solar_resource])}, itemStyle: {{color:'#ef4444'}}, lineStyle: {{width:2}} }}
  ]
}});

// ==================== CHART 7 ====================
var c7 = echarts.init(document.getElementById('chart7'));
var sun7 = {json.dumps([s['name'] for s in stations if s['sur_m'] > 0], ensure_ascii=False)};
c7.setOption({{
  title: {{ text: '各场站月场用电率（%）', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  grid: {{ top: 40, bottom: 80, left: 50, right: 30 }},
  xAxis: {{ type: 'category', data: sun7, axisLabel: {{ rotate: 55, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: '%' }},
  series: [{{
    type: 'bar',
    data: {json.dumps([round(s['sur_m']*100,2) for s in stations if s['sur_m'] > 0])}.map(function(v){{ return {{ value: v, itemStyle: {{ color: v > 3 ? '#ef4444' : v > 1 ? '#f59e0b' : '#10b981' }} }}; }}),
    label: {{ show: true, position: 'top', fontSize: 10, formatter: '{{c}}%' }}
  }}]
}});

// ==================== CHART 8 (FIXED) ====================
var c8 = echarts.init(document.getElementById('chart8'));
var gh_names = {json.dumps([s['name'] for s in sorted(stations, key=lambda x: -(x['gen_cum']/x['capacity'] if x['capacity']>0 else 0))[:15]], ensure_ascii=False)};
var gh_vals = {json.dumps([round(s['gen_cum']/s['capacity'],1) if s['capacity']>0 else 0 for s in sorted(stations, key=lambda x: -(x['gen_cum']/x['capacity'] if x['capacity']>0 else 0))[:15]])};
c8.setOption({{
  title: {{ text: '各场站累计发电小时数 TOP15（h）', left: 'center', textStyle: {{fontSize:14}} }},
  tooltip: {{ trigger: 'axis' }},
  grid: {{ top: 40, bottom: 80, left: 60, right: 30 }},
  xAxis: {{ type: 'category', data: gh_names, axisLabel: {{ rotate: 55, fontSize: 11 }} }},
  yAxis: {{ type: 'value', name: 'h' }},
  series: [{{
    type: 'bar', data: gh_vals,
    itemStyle: {{ color: '#3b82f6' }},
    label: {{ show: true, position: 'top', fontSize: 10 }}
  }}]
}});

window.addEventListener('resize', function() {{
  [c1,c2,c3a,c3b,c4,c5,c6a,c6b,c7,c8].forEach(function(c){{ c.resize(); }});
}});
</script>
</body>
</html>'''

with open('山东公司2026年4月电量生产分析报告.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("OK - report regenerated")
