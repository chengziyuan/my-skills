---
name: operations-ledger-analyst
description: 分析运营中心部门事项报备台账（Excel格式），自动生成包含统计图表、关键洞察和事项明细的 HTML + PDF 分析报告。当用户上传任何名称含"台账"、"报备"、"事项"的 Excel 文件，或说"帮我分析台账"、"生成台账报告"、"出一份事项分析"时，立即触发此 Skill。即使用户只说"分析一下"并附上 xlsx 文件，也应优先考虑使用此 Skill。
---

# 运营中心事项报备台账分析 Skill

## 功能概述
读取运营中心部门事项报备台账 Excel 文件，自动完成：
1. 数据解析（含混合日期格式处理）
2. 完成状态判断（已完成 / 进行中 / 逾期）
3. 生成5张分析图表
4. 输出带排序分组的事项明细表
5. 同时输出 HTML 和 PDF 两个版本

---

## 文件结构约定

台账 Excel 通常含多个 Sheet，**只分析"当年台账"Sheet**（如"2026年台账"）。

| 关键列 | 说明 |
|--------|------|
| 二级部门名称 | 部门 |
| 事项主题 | 事项名称 |
| 事项类型 | "重大事项" 或 "一般事项" |
| 预计完成时间 | 混合格式（见日期解析） |
| 实际完成时间 | 混合格式，空值=未完成 |
| 填报人 | 人员姓名 |
| 2.变更原因（外规）/ Unnamed:14 | 外规修订勾选列 |
| 3.变更原因（业务部门）/ Unnamed:19/20 | 业务驱动勾选列 |
| 4.变更原因（本部门）/ Unnamed:27 | 本部门驱动勾选列 |

读取方式：`header=0, skiprows=[1]`（第0行是列名，第1行是子标题，跳过）

---

## 日期解析（重要）

Excel 中日期存在三种混合格式，必须全部处理：

```python
import re, pandas as pd

def parse_date(val):
    if pd.isna(val): return pd.NaT
    # 数字序列号（如 46038 = 2026-01-16）
    if isinstance(val, (int, float)):
        try: return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(val))
        except: return pd.NaT
    s = str(val).strip().split('\n')[0]  # 去掉换行后备注
    # 仅年月格式（如 "2026.3" → 当月1日）
    m = re.match(r'^(\d{4})\.(\d{1,2})$', s)
    if m: return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    # 标准字符串格式
    for fmt in ['%Y.%m.%d', '%Y/%m/%d', '%Y-%m-%d']:
        try: return pd.to_datetime(s, format=fmt)
        except: continue
    return pd.NaT

def fmt_date(val):
    """原始值 → 可读字符串，失败返回原始字符串"""
    dt = parse_date(val)
    if pd.notna(dt): return dt.strftime('%Y-%m-%d')
    if pd.isna(val): return '—'
    return str(val).strip().split('\n')[0]
```

---

## 状态判断逻辑

```python
df['已完成'] = df['实际完成_dt'].notna()
df['逾期']   = df.apply(lambda r: r['实际完成_dt'] > r['预计完成_dt']
                         if pd.notna(r['实际完成_dt']) and pd.notna(r['预计完成_dt'])
                         else False, axis=1)
df['进行中'] = df['实际完成_dt'].isna() & df['预计完成_dt'].notna()
```

---

## 图表清单（共5张）

| # | 图表 | 类型 | 说明 |
|---|------|------|------|
| 1 | 各部门事项完成情况 | 横向堆叠柱状图 | 已完成/进行中/逾期分色，末尾显示总数 |
| 2 | 重大/一般事项分布 | 横向堆叠柱状图 | 重大(红)/一般(蓝) |
| 3 | 月度计划 vs 实际完成 | 分组柱状图 | 过滤异常年份（只保留当年+次年） |
| 4 | 变更原因分布 | 饼图 | 外规修订/业务驱动/本部门驱动/未标注 |
| 5 | 填报人事项数 Top10 | 竖向柱状图 | 柱顶显示数值 |

颜色规范：主色 `#4F46E5`，成功 `#10B981`，警告 `#F59E0B`，危险 `#EF4444`，紫 `#8B5CF6`

中文字体设置：
```python
import matplotlib.font_manager as fm
font_path = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'
zh = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = zh.get_name()
# seaborn heatmap 等需手动设置每个 tick
for tick in ax.get_xticklabels(): tick.set_fontproperties(zh)
```

---

## 明细表排序规则

```python
def sort_key(row):
    if row['事项类型'] == '重大事项': return 1   # 最优先，不论完成与否
    elif row['逾期']:                  return 2   # 逾期一般事项
    elif not row['已完成']:            return 3   # 进行中一般事项
    else:                              return 4   # 已完成一般事项

df['排序键'] = df.apply(sort_key, axis=1)
df = df.sort_values(['排序键', '二级部门名称']).reset_index(drop=True)
```

分组标题行配色：

| 分组 | 标题文字 | 背景色 | 文字色 |
|------|----------|--------|--------|
| 重大事项 | 🔴 重大事项 | `#FFF1F2` | `#DC2626` |
| 逾期一般 | ⏰ 逾期一般事项 | `#FFF7ED` | `#D97706` |
| 进行中一般 | 🔵 进行中一般事项 | `#EFF6FF` | `#2563EB` |
| 已完成一般 | ✅ 已完成一般事项 | `#F0FDF4` | `#16A34A` |

**特别规则**：进行中一般事项组内，按部门插入二级子标题行：
```html
<tr><td colspan="7" style="background:#DBEAFE;color:#1D4ED8;font-weight:600;
    font-size:11px;padding:6px 20px;">　　📁 {部门名称}</td></tr>
```

行背景色：重大=`#FFF1F2`，逾期=`#FEF2F2`，进行中=`#EFF6FF`，已完成=`#F0FDF4`

---

## 顶部指标卡（5个）

| 图标 | 标签 | 值 | 颜色 |
|------|------|----|------|
| 📋 | 总事项数 | 总行数 | `#4F46E5` |
| ✅ | 已完成 | 数量(完成率%) | `#10B981` |
| 🔵 | 进行中 | 数量 | `#3B82F6` |
| ⏰ | 逾期 | 数量 | `#EF4444` |
| 🔴 | 重大事项 | 数量 | `#8B5CF6` |

---

## 关键洞察（自动生成6条）

1. 总事项/已完成/完成率/进行中
2. 重大事项数量及占比
3. 逾期数量及涉及部门
4. 事项最多的部门及占比
5. 填报最活跃人员及数量
6. 变更原因主导类型

---

## 输出规范

```python
# 文件命名
html_path = f'/mnt/user-data/outputs/运营中心事项台账分析_{year}.html'
pdf_path  = f'/mnt/user-data/outputs/运营中心事项台账分析_{year}.pdf'

# HTML → PDF 转换
from weasyprint import HTML
HTML(filename=html_path).write_pdf(pdf_path)

# 同时用 present_files 提供两个文件
```

在对话中简短说明：完成率、逾期情况、最需关注的部门。

---

## 安装依赖

```bash
pip install pandas openpyxl matplotlib seaborn weasyprint --break-system-packages -q
```

---

## 常见问题

| 问题 | 处理方式 |
|------|----------|
| 日期显示为数字（如46038） | 用 `fmt_date()` 转换，勿直接显示原始值 |
| "2026.3" 被解析为1905年 | 用正则匹配年月格式，转为当月1日 |
| seaborn heatmap 中文乱码 | 手动对每个 tick 调用 `set_fontproperties(zh)` |
| Sheet名称不固定 | 优先选含"当年"或最大年份的Sheet |
| 编码错误 | 尝试 gbk、gb2312 |
