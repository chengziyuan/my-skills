"""
运营中心事项报备台账分析脚本
用法：python analyze.py <excel文件路径> [sheet名称]
输出：HTML + PDF 报告到 /mnt/user-data/outputs/
"""
import sys, re, os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import base64
from io import BytesIO
from datetime import datetime

# ── 字体 ──────────────────────────────────────────────────
font_path = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'
zh = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = zh.get_name()
plt.rcParams['axes.unicode_minus'] = False

COLORS = ['#4F46E5','#10B981','#F59E0B','#EF4444','#8B5CF6','#06B6D4']

# ── 日期处理 ──────────────────────────────────────────────
def parse_date(val):
    if pd.isna(val): return pd.NaT
    if isinstance(val, (int, float)):
        try: return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(val))
        except: return pd.NaT
    s = str(val).strip().split('\n')[0]
    m = re.match(r'^(\d{4})\.(\d{1,2})$', s)
    if m: return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    for fmt in ['%Y.%m.%d','%Y/%m/%d','%Y-%m-%d']:
        try: return pd.to_datetime(s, format=fmt)
        except: continue
    return pd.NaT

def fmt_date(val):
    dt = parse_date(val)
    if pd.notna(dt): return dt.strftime('%Y-%m-%d')
    if pd.isna(val): return '—'
    return str(val).strip().split('\n')[0]

# ── 图表工具 ──────────────────────────────────────────────
def fig_to_b64(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=140, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode()
    plt.close(fig)
    return b64

def analyze(file_path, sheet_name=None):
    # ── 选Sheet ──────────────────────────────────────────
    xls = pd.ExcelFile(file_path)
    if sheet_name is None:
        # 自动选含最大年份的台账Sheet
        candidates = [s for s in xls.sheet_names if '台账' in s]
        sheet_name = candidates[-1] if candidates else xls.sheet_names[0]
    year = re.search(r'\d{4}', sheet_name)
    year = year.group() if year else datetime.now().strftime('%Y')

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, skiprows=[1])

    # ── 状态计算 ─────────────────────────────────────────
    df['预计完成_dt'] = df['预计完成时间'].apply(parse_date)
    df['实际完成_dt'] = df['实际完成时间'].apply(parse_date)
    df['已完成'] = df['实际完成_dt'].notna()
    df['逾期']   = df.apply(lambda r: r['实际完成_dt'] > r['预计完成_dt']
                             if pd.notna(r['实际完成_dt']) and pd.notna(r['预计完成_dt']) else False, axis=1)
    df['进行中'] = df['实际完成_dt'].isna() & df['预计完成_dt'].notna()

    # ── 变更原因 ─────────────────────────────────────────
    def has_check(val): return str(val).strip() == '√'
    df['变更_外规修订']   = df.get('2.变更原因（外规）', pd.Series()).apply(has_check) | df.get('Unnamed: 14', pd.Series()).apply(has_check)
    df['变更_业务驱动']   = df.get('3.变更原因（业务部门）', pd.Series()).apply(has_check) | df.get('Unnamed: 19', pd.Series()).apply(has_check) | df.get('Unnamed: 20', pd.Series()).apply(has_check)
    df['变更_本部门驱动'] = df.get('4.变更原因（本部门）', pd.Series()).apply(has_check) | df.get('Unnamed: 27', pd.Series()).apply(has_check)

    # ── 排序 ─────────────────────────────────────────────
    def sort_key(row):
        if row['事项类型'] == '重大事项': return 1
        elif row['逾期']:                  return 2
        elif not row['已完成']:            return 3
        else:                              return 4
    df['排序键'] = df.apply(sort_key, axis=1)
    df = df.sort_values(['排序键','二级部门名称']).reset_index(drop=True)

    # ── 图表 ─────────────────────────────────────────────
    charts = []

    # 图1：各部门完成情况
    dept_g = df.groupby('二级部门名称').agg(已完成=('已完成','sum'),进行中=('进行中','sum'),逾期=('逾期','sum')).reset_index().sort_values('已完成',ascending=True)
    fig, ax = plt.subplots(figsize=(8,5))
    y = range(len(dept_g))
    ax.barh(list(y), dept_g['已完成'], color='#10B981', label='已完成')
    ax.barh(list(y), dept_g['进行中'], left=dept_g['已完成'], color='#4F46E5', label='进行中')
    ax.barh(list(y), dept_g['逾期'],   left=dept_g['已完成']+dept_g['进行中'], color='#EF4444', label='逾期')
    ax.set_yticks(list(y)); ax.set_yticklabels(dept_g['二级部门名称'].tolist(), fontproperties=zh)
    ax.set_xlabel('事项数量', fontproperties=zh)
    ax.set_title('各部门事项完成情况', fontsize=14, fontweight='bold', fontproperties=zh, pad=12)
    ax.legend(prop=zh, loc='lower right')
    for i, row in enumerate(dept_g.itertuples()):
        total = row.已完成 + row.进行中 + row.逾期
        ax.text(total+0.3, i, str(total), va='center', fontsize=9)
    fig.tight_layout(); charts.append(('各部门事项完成情况', fig_to_b64(fig)))

    # 图2：重大/一般
    type_g = df.groupby(['二级部门名称','事项类型']).size().unstack(fill_value=0).reset_index().sort_values('一般事项',ascending=True)
    fig, ax = plt.subplots(figsize=(8,5))
    y = range(len(type_g))
    if '重大事项' in type_g.columns:
        ax.barh(list(y), type_g['重大事项'], color='#EF4444', label='重大事项')
        ax.barh(list(y), type_g['一般事项'], left=type_g['重大事项'], color='#4F46E5', label='一般事项')
    else:
        ax.barh(list(y), type_g['一般事项'], color='#4F46E5', label='一般事项')
    ax.set_yticks(list(y)); ax.set_yticklabels(type_g['二级部门名称'].tolist(), fontproperties=zh)
    ax.set_xlabel('事项数量', fontproperties=zh)
    ax.set_title('各部门重大/一般事项分布', fontsize=14, fontweight='bold', fontproperties=zh, pad=12)
    ax.legend(prop=zh); fig.tight_layout(); charts.append(('重大/一般事项分布', fig_to_b64(fig)))

    # 图3：月度对比
    monthly_done = df[df['已完成']]['实际完成_dt'].dt.to_period('M').value_counts().sort_index()
    monthly_plan = df['预计完成_dt'].dropna()
    monthly_plan = monthly_plan[monthly_plan.dt.year.between(int(year), int(year)+1)].dt.to_period('M').value_counts().sort_index()
    months = sorted(set(list(monthly_done.index)+list(monthly_plan.index)))
    fig, ax = plt.subplots(figsize=(9,4.5))
    x = range(len(months))
    ax.bar([i-0.2 for i in x], [monthly_plan.get(m,0) for m in months], width=0.38, color='#C7D2FE', label='计划事项数')
    ax.bar([i+0.2 for i in x], [monthly_done.get(m,0) for m in months], width=0.38, color='#4F46E5', label='实际完成数')
    ax.set_xticks(list(x)); ax.set_xticklabels([str(m) for m in months], fontproperties=zh, rotation=15)
    ax.set_ylabel('事项数量', fontproperties=zh)
    ax.set_title('月度计划 vs 实际完成对比', fontsize=14, fontweight='bold', fontproperties=zh, pad=12)
    ax.legend(prop=zh); fig.tight_layout(); charts.append(('月度计划 vs 实际完成', fig_to_b64(fig)))

    # 图4：变更原因饼图
    change_data = {'外规修订驱动': int(df['变更_外规修订'].sum()), '业务部门驱动': int(df['变更_业务驱动'].sum()), '本部门驱动': int(df['变更_本部门驱动'].sum())}
    none_count = int((~df['变更_外规修订'] & ~df['变更_业务驱动'] & ~df['变更_本部门驱动']).sum())
    if none_count > 0: change_data['未标注原因'] = none_count
    fig, ax = plt.subplots(figsize=(6,5))
    wedges, texts, autotexts = ax.pie(list(change_data.values()), labels=list(change_data.keys()), autopct='%1.1f%%',
        colors=['#4F46E5','#10B981','#F59E0B','#94A3B8'][:len(change_data)], startangle=140, textprops={'fontproperties': zh})
    for at in autotexts: at.set_fontsize(9)
    ax.set_title('事项变更原因分布', fontsize=14, fontweight='bold', fontproperties=zh)
    fig.tight_layout(); charts.append(('变更原因分布', fig_to_b64(fig)))

    # 图5：填报人 Top10
    reporter_g = df['填报人'].value_counts().head(10)
    fig, ax = plt.subplots(figsize=(8,4.5))
    bars = ax.bar(reporter_g.index.tolist(), reporter_g.values, color=COLORS[:len(reporter_g)])
    ax.set_ylabel('事项数量', fontproperties=zh)
    ax.set_title('填报人事项数量 Top10', fontsize=14, fontweight='bold', fontproperties=zh, pad=12)
    for tick in ax.get_xticklabels(): tick.set_fontproperties(zh)
    for b in bars: ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.1, str(int(b.get_height())), ha='center', va='bottom', fontsize=9)
    fig.tight_layout(); charts.append(('填报人事项数量 Top10', fig_to_b64(fig)))

    # ── 指标 ─────────────────────────────────────────────
    total = len(df); done = int(df['已完成'].sum()); wip = int(df['进行中'].sum())
    overdue = int(df['逾期'].sum()); major = int((df['事项类型']=='重大事项').sum())
    completion_rate = round(done/total*100, 1)
    top_dept = df.groupby('二级部门名称').size().idxmax()
    busiest_reporter = df['填报人'].value_counts().idxmax()
    busiest_reporter_cnt = int(df['填报人'].value_counts().iloc[0])
    overdue_depts = df[df['逾期']]['二级部门名称'].value_counts()
    overdue_str = "、".join([f"{d}({n}项)" for d,n in overdue_depts.items()]) if len(overdue_depts) else "无"

    insights = [
        f"📋 台账共 <b>{total}</b> 项事项，已完成 <b>{done}</b> 项，整体完成率 <b>{completion_rate}%</b>，进行中 <b>{wip}</b> 项",
        f"🔴 重大事项 <b>{major}</b> 项，占比 {round(major/total*100,1)}%，需重点关注进展",
        f"⏰ 当前逾期事项 <b>{overdue}</b> 项，涉及部门：{overdue_str}",
        f"🏢 <b>{top_dept}</b> 报备事项最多，承担了全部事项的 {round(df.groupby('二级部门名称').size()[top_dept]/total*100,1)}%",
        f"✍️ 填报最活跃人员为 <b>{busiest_reporter}</b>，共填报 {busiest_reporter_cnt} 项",
        f"📈 业务部门驱动变更占主导（{change_data.get('业务部门驱动',0)}项），说明业务发展仍是运营中心工作的主要推动力",
    ]

    # ── 明细表 ───────────────────────────────────────────
    group_labels = {
        1: ('🔴 重大事项',       '#FFF1F2','#DC2626'),
        2: ('⏰ 逾期一般事项',   '#FFF7ED','#D97706'),
        3: ('🔵 进行中一般事项', '#EFF6FF','#2563EB'),
        4: ('✅ 已完成一般事项', '#F0FDF4','#16A34A'),
    }
    table_rows = ""; prev_key = None; prev_dept = None
    for _, row in df.iterrows():
        key = int(row['排序键']); dept = row['二级部门名称']
        if key != prev_key:
            label, bg, color = group_labels[key]
            table_rows += f'<tr><td colspan="7" style="background:{bg};color:{color};font-weight:700;font-size:12px;padding:8px 12px;border-bottom:2px solid {color}20;border-top:2px solid {color}40;">{label}</td></tr>'
            prev_key = key; prev_dept = None
        if key == 3 and dept != prev_dept:
            table_rows += f'<tr><td colspan="7" style="background:#DBEAFE;color:#1D4ED8;font-weight:600;font-size:11px;padding:6px 20px;border-bottom:1px solid #BFDBFE;">　　📁 {dept}</td></tr>'
            prev_dept = dept
        status = '🔴 逾期' if row['逾期'] else ('✅ 已完成' if row['已完成'] else '🔵 进行中')
        sc = 'major' if row['事项类型']=='重大事项' else ('overdue' if row['逾期'] else ('done' if row['已完成'] else 'wip'))
        table_rows += f"""<tr class="{sc}">
            <td>{dept}</td><td style="max-width:280px;word-break:break-all">{row['事项主题']}</td>
            <td><span class="tag {'tag-major' if row['事项类型']=='重大事项' else 'tag-normal'}">{row['事项类型']}</span></td>
            <td style="white-space:nowrap">{fmt_date(row['预计完成时间'])}</td>
            <td style="white-space:nowrap">{fmt_date(row['实际完成时间'])}</td>
            <td>{row['填报人'] if not pd.isna(row['填报人']) else '—'}</td>
            <td style="white-space:nowrap">{status}</td></tr>"""

    # ── HTML 组装 ─────────────────────────────────────────
    cards_html = "".join([f'<div class="card" style="border-top:4px solid {c}"><div class="card-icon">{ic}</div><div class="card-val" style="color:{c}">{v}</div><div class="card-label">{lb}</div></div>'
        for ic,lb,v,c in [('📋','总事项数',str(total),'#4F46E5'),('✅','已完成',f'{done} ({completion_rate}%)','#10B981'),
                           ('🔵','进行中',str(wip),'#3B82F6'),('⏰','逾期',str(overdue),'#EF4444'),('🔴','重大事项',str(major),'#8B5CF6')]])
    charts_section = "".join(f'<div class="chart-card"><h3 class="chart-title">{t}</h3><img src="data:image/png;base64,{b}" style="width:100%;border-radius:6px;"></div>' for t,b in charts)
    insights_html = "".join(f'<li>{i}</li>' for i in insights)

    html = f"""<!DOCTYPE html><html lang="zh"><head><meta charset="UTF-8">
<title>运营中心事项台账分析 - {year}年</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;background:#F1F5F9;color:#1E293B;font-size:14px}}
.header{{background:linear-gradient(135deg,#1E40AF,#4F46E5);color:white;padding:32px 40px}}
.header h1{{font-size:22px;font-weight:700}}.header .sub{{margin-top:6px;opacity:.8;font-size:13px}}
.container{{max-width:1200px;margin:0 auto;padding:28px 24px}}.section{{margin-bottom:32px}}
.section-title{{font-size:16px;font-weight:700;margin-bottom:14px;padding-left:10px;border-left:4px solid #4F46E5}}
.cards{{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:28px}}
.card{{background:white;border-radius:12px;padding:18px 16px;text-align:center;box-shadow:0 1px 6px rgba(0,0,0,.07)}}
.card-icon{{font-size:22px;margin-bottom:6px}}.card-val{{font-size:22px;font-weight:700;margin-bottom:4px}}.card-label{{font-size:12px;color:#64748B}}
.insights{{background:#EEF2FF;border-radius:12px;padding:18px 22px}}
.insights ul{{list-style:none}}.insights li{{padding:8px 0;font-size:13px;line-height:1.6;border-bottom:1px solid #C7D2FE}}.insights li:last-child{{border-bottom:none}}
.charts-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(440px,1fr));gap:18px}}
.chart-card{{background:white;border-radius:12px;padding:18px;box-shadow:0 1px 6px rgba(0,0,0,.07)}}
.chart-title{{font-size:13px;font-weight:600;color:#4F46E5;margin-bottom:10px}}
.table-wrap{{overflow-x:auto;border-radius:12px;box-shadow:0 1px 6px rgba(0,0,0,.07)}}
table.detail{{width:100%;border-collapse:collapse;background:white;font-size:12px}}
table.detail th{{background:#1E293B;color:white;padding:10px 12px;text-align:left;white-space:nowrap}}
table.detail td{{padding:8px 12px;border-bottom:1px solid #F1F5F9}}
table.detail tr.major td{{background:#FFF1F2}}table.detail tr.done td{{background:#F0FDF4}}
table.detail tr.overdue td{{background:#FEF2F2}}table.detail tr.wip td{{background:#EFF6FF}}
table.detail tr:hover td{{filter:brightness(.97)}}
.tag{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}}
.tag-major{{background:#FEE2E2;color:#DC2626}}.tag-normal{{background:#E0E7FF;color:#4338CA}}
</style></head><body>
<div class="header"><h1>📊 运营中心部门事项报备台账分析</h1>
<div class="sub">{year}年台账 &nbsp;|&nbsp; 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')} &nbsp;|&nbsp; 数据共 {total} 条</div></div>
<div class="container">
<div class="cards">{cards_html}</div>
<div class="section"><div class="section-title">💡 关键洞察</div><div class="insights"><ul>{insights_html}</ul></div></div>
<div class="section"><div class="section-title">📈 数据可视化</div><div class="charts-grid">{charts_section}</div></div>
<div class="section"><div class="section-title">📋 事项明细</div>
<div class="table-wrap"><table class="detail">
<tr><th>部门</th><th>事项主题</th><th>类型</th><th>预计完成</th><th>实际完成</th><th>填报人</th><th>状态</th></tr>
{table_rows}</table></div></div>
</div></body></html>"""

    # ── 输出 ─────────────────────────────────────────────
    os.makedirs('/mnt/user-data/outputs', exist_ok=True)
    html_path = f'/mnt/user-data/outputs/运营中心事项台账分析_{year}.html'
    pdf_path  = f'/mnt/user-data/outputs/运营中心事项台账分析_{year}.pdf'

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    from weasyprint import HTML
    HTML(filename=html_path).write_pdf(pdf_path)

    print(f"✅ 报告已生成：{html_path}")
    print(f"✅ PDF已生成：{pdf_path}")
    return html_path, pdf_path, {'total':total,'done':done,'completion_rate':completion_rate,'overdue':overdue,'overdue_str':overdue_str,'top_dept':top_dept}

if __name__ == '__main__':
    file_path = sys.argv[1] if len(sys.argv) > 1 else None
    sheet = sys.argv[2] if len(sys.argv) > 2 else None
    if not file_path:
        print("用法：python analyze.py <excel文件路径> [sheet名称]")
        sys.exit(1)
    html_path, pdf_path, summary = analyze(file_path, sheet)
    print(f"\n📊 摘要：共{summary['total']}项，完成率{summary['completion_rate']}%，逾期{summary['overdue']}项（{summary['overdue_str']}），事项最多部门：{summary['top_dept']}")
