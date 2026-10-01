# -*- coding: utf-8 -*-
"""
中美国债收益率 抓取 + 表格 + 图表面板
========================================
数据来源：东方财富数据中心-中美国债收益率
  页面  : https://data.eastmoney.com/cjsj/zmgzsyl.html
  接口  : https://datacenter-web.eastmoney.com/api/data/v1/get (reportName=RPTA_WEB_TREASURYYIELD)

功能：
  1) 从东方财富接口抓取全量中美国债收益率数据
  2) 输出 CSV 表格 和 带样式的 Excel 表格
  3) 生成一个 UI 漂亮的 HTML 数据面板（ECharts 图表 + KPI 卡片 + 数据表格）

用法：
  python fetch_treasury.py            # 使用默认输出目录（脚本所在目录）
  python fetch_treasury.py D:/out     # 指定输出目录

依赖：
  pip install requests pandas openpyxl
"""

import argparse
import json
import os
import sys
import time

import pandas as pd
import requests

# ----------------------------- 数据源配置 -----------------------------
API_URL = "https://datacenter-web.eastmoney.com/api/data/v1/get"
REPORT_NAME = "RPTA_WEB_TREASURYYIELD"

# 字段 key -> 中文列名（列顺序即输出顺序）
FIELD_MAP = [
    ("SOLAR_DATE", "日期"),
    ("EMM00588704", "中国国债收益率:2年"),
    ("EMM00166462", "中国国债收益率:5年"),
    ("EMM00166466", "中国国债收益率:10年"),
    ("EMM00166469", "中国国债收益率:30年"),
    ("EMM01276014", "中国国债收益率:10年-2年"),
    ("EMM00000024", "中国GDP年增率(%)"),
    ("EMG00001306", "美国国债收益率:2年"),
    ("EMG00001308", "美国国债收益率:5年"),
    ("EMG00001310", "美国国债收益率:10年"),
    ("EMG00001312", "美国国债收益率:30年"),
    ("EMG01339436", "美国国债收益率:10年-2年"),
    ("EMG00159635", "美国GDP年增率(%)"),
]

PAGE_SIZE = 2000  # 每页条数
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Referer": "https://data.eastmoney.com/",
}


def fetch_all_rows() -> list:
    """循环分页抓取全部数据，返回原始字典列表。"""
    rows, page = [], 1
    while True:
        params = {
            "reportName": REPORT_NAME,
            "columns": "ALL",
            "sortColumns": "SOLAR_DATE",
            "sortTypes": "-1",
            "pageSize": PAGE_SIZE,
            "pageNumber": page,
            "source": "WEB",
            "client": "WEB",
        }
        for attempt in range(3):  # 简单重试
            try:
                r = requests.get(API_URL, params=params, headers=HEADERS, timeout=30)
                r.raise_for_status()
                data = r.json()
                break
            except Exception as e:  # noqa: BLE001
                if attempt == 2:
                    raise RuntimeError(f"抓取第 {page} 页失败: {e}") from e
                time.sleep(1.5)

        result = (data.get("result") or {})
        batch = result.get("data") or []
        if not batch:
            break
        rows.extend(batch)
        total = result.get("count")
        pages = result.get("pages") or 1
        print(f"[抓取] 第 {page}/{pages} 页，累计 {len(rows)}/{total} 条")
        if page >= pages:
            break
        page += 1
        time.sleep(0.5)
    return rows


def build_dataframe(rows: list) -> pd.DataFrame:
    """把原始接口数据整理成 DataFrame，并统一列名、处理空值。"""
    records = []
    for item in rows:
        rec = {}
        for key, name in FIELD_MAP:
            val = item.get(key)
            # 接口空值可能是 None 或空串，统一转成 NaN，并保留最多 4 位小数
            if val in (None, ""):
                rec[name] = None
            elif isinstance(val, (int, float)):
                rec[name] = round(float(val), 4)
            else:
                rec[name] = val
        records.append(rec)
    df = pd.DataFrame(records, columns=[n for _, n in FIELD_MAP])
    df["日期"] = pd.to_datetime(df["日期"]).dt.strftime("%Y-%m-%d")
    return df


def save_table(df: pd.DataFrame, out_dir: str) -> tuple:
    """输出 CSV 与带样式的 Excel 表格，返回路径。"""
    csv_path = os.path.join(out_dir, "中美国债收益率.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    xlsx_path = os.path.join(out_dir, "中美国债收益率.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="中美国债收益率", index=False)
        ws = writer.sheets["中美国债收益率"]

        # 表头样式
        from openpyxl.styles import Alignment, Font, PatternFill
        head_fill = PatternFill("solid", fgColor="1F3864")
        for cell in ws[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # 列宽
        widths = {"A": 12}
        for idx, col in enumerate(df.columns, start=2):
            widths[chr(ord("A") + idx - 1)] = max(18, min(24, len(col) * 2 + 4))
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w

        # 冻结首行
        ws.freeze_panes = "A2"
    return csv_path, xlsx_path


# ----------------------------- 图表面板 -----------------------------
def render_dashboard(df: pd.DataFrame, out_dir: str) -> str:
    """用 ECharts 生成一个 UI 漂亮的 HTML 数据面板，返回路径。"""
    # 只保留两列都有数据的日期，避免断点过多；再补一段口径说明
    data = df.dropna(subset=["中国国债收益率:10年", "美国国债收益率:10年"])

    latest = df.iloc[0]  # 数据已按日期倒序

    def v(col):
        val = latest.get(col)
        return None if pd.isna(val) else round(float(val), 2)

    cn_10 = v("中国国债收益率:10年")
    us_10 = v("美国国债收益率:10年")
    spread = round(cn_10 - us_10, 2) if cn_10 is not None and us_10 is not None else None
    latest_date = str(latest["日期"])

    def series(col, label):
        s = data[["日期", col]].dropna()
        return [[d, round(float(y), 3)] for d, y in zip(s["日期"], s[col])]

    chart1 = {  # 中国 vs 美国 10年期收益率
        "cn10": series("中国国债收益率:10年", "中国10年"),
        "us10": series("美国国债收益率:10年", "美国10年"),
    }
    chart2 = {  # 各期限收益率
        "cn2": series("中国国债收益率:2年", "中国2年"),
        "cn5": series("中国国债收益率:5年", "中国5年"),
        "cn10": series("中国国债收益率:10年", "中国10年"),
        "cn30": series("中国国债收益率:30年", "中国30年"),
        "us2": series("美国国债收益率:2年", "美国2年"),
        "us5": series("美国国债收益率:5年", "美国5年"),
        "us10": series("美国国债收益率:10年", "美国10年"),
        "us30": series("美国国债收益率:30年", "美国30年"),
    }
    chart3 = {  # 10年-2年期限利差
        "cn": series("中国国债收益率:10年-2年", "中国10年-2年"),
        "us": series("美国国债收益率:10年-2年", "美国10年-2年"),
    }

    # 表格预览：最近 30 条
    preview = data.head(30)
    table_rows = []
    for _, row in preview.iterrows():
        cells = []
        for key, _name in FIELD_MAP:
            val = row.get(_name)
            cells.append("" if pd.isna(val) else (str(val) if not isinstance(val, float) else f"{val:.4f}".rstrip("0").rstrip(".")))
        table_rows.append(cells)
    table_head = [n for _, n in FIELD_MAP]

    payload = {
        "latest_date": latest_date,
        "cn_10": cn_10, "us_10": us_10, "spread": spread,
        "cn_2": v("中国国债收益率:2年"), "cn_30": v("中国国债收益率:30年"),
        "us_2": v("美国国债收益率:2年"), "us_30": v("美国国债收益率:30年"),
        "chart1": chart1, "chart2": chart2, "chart3": chart3,
        "table_head": table_head, "table_rows": table_rows,
        "total_rows": int(len(data)),
    }
    html = (DASHBOARD_HTML
            .replace("__DATA__", json.dumps(payload, ensure_ascii=False))
            .replace("__TOTAL__", str(int(len(data))))
            .replace("__LATEST__", latest_date)
            .replace("__CN10__", _fmt(cn_10))
            .replace("__US10__", _fmt(us_10))
            .replace("__SPREAD__", _fmt(spread))
            .replace("__CN2__", _fmt(v("中国国债收益率:2年")))
            .replace("__CN30__", _fmt(v("中国国债收益率:30年")))
            .replace("__US2__", _fmt(v("美国国债收益率:2年")))
            .replace("__US30__", _fmt(v("美国国债收益率:30年"))))
    html_path = os.path.join(out_dir, "中美国债收益率_数据面板.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    return html_path


def _fmt(x):
    return "—" if x is None else f"{x:.2f}"


DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>中美国债收益率数据面板</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
<style>
  :root{
    --bg:#0f172a; --card:#1e293b; --card2:#111c34; --line:#334155;
    --txt:#e2e8f0; --sub:#94a3b8; --cn:#38bdf8; --us:#fb923c;
    --cn2:#22d3ee; --us2:#fbbf24; --accent:#6366f1; --up:#34d399; --down:#f87171;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:radial-gradient(1200px 600px at 20% -10%,#1e3a5f 0%,var(--bg) 55%);color:var(--txt);
    font-family:"Segoe UI","Microsoft YaHei",system-ui,sans-serif;min-height:100vh;padding:24px 28px 40px}
  .wrap{max-width:1240px;margin:0 auto}
  header{display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:12px;margin-bottom:22px}
  h1{font-size:26px;font-weight:700;letter-spacing:.5px}
  h1 small{display:block;font-size:13px;font-weight:400;color:var(--sub);margin-top:6px}
  .badge{background:#1e40af33;border:1px solid #2563eb55;color:#93c5fd;padding:6px 12px;border-radius:999px;font-size:12.5px}
  .kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:22px}
  .kpi{background:linear-gradient(160deg,var(--card),var(--card2));border:1px solid var(--line);border-radius:16px;padding:16px 18px;position:relative;overflow:hidden}
  .kpi::after{content:"";position:absolute;inset:0;background:radial-gradient(120px 80px at 100% 0%,rgba(99,102,241,.18),transparent)}
  .kpi .lb{font-size:12px;color:var(--sub);margin-bottom:8px}
  .kpi .num{font-size:26px;font-weight:700;font-variant-numeric:tabular-nums}
  .kpi .un{font-size:12px;color:var(--sub);margin-left:3px;font-weight:400}
  .kpi .tag{font-size:11px;margin-top:7px;color:var(--sub)}
  .cn .num{color:var(--cn)} .us .num{color:var(--us)}
  .spread .num{color:var(--accent)} .date .num{font-size:17px;font-weight:600}
  .grid{display:grid;gap:16px;margin-bottom:16px}
  .g1{grid-template-columns:3fr 2fr}
  .g2{grid-template-columns:1fr 1fr}
  .card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:16px 18px}
  .card h2{font-size:15px;font-weight:600;margin-bottom:4px;display:flex;align-items:center;gap:8px}
  .card .sub{font-size:12px;color:var(--sub);margin-bottom:10px}
  .chart{width:100%;height:340px}
  .legend{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;color:var(--sub);margin-bottom:8px}
  .legend i{display:inline-block;width:14px;height:3px;border-radius:2px;margin-right:5px;vertical-align:middle}
  table{width:100%;border-collapse:collapse;font-size:12.5px;overflow:hidden}
  thead th{background:#0b1526;color:var(--sub);font-weight:600;padding:9px 10px;text-align:right;white-space:nowrap;position:sticky;top:0}
  thead th:first-child{text-align:left}
  tbody td{padding:7px 10px;text-align:right;border-top:1px solid #16233c;white-space:nowrap;font-variant-numeric:tabular-nums}
  tbody td:first-child{text-align:left;color:var(--sub)}
  tbody tr:hover td{background:#17233d}
  .scroll{max-height:420px;overflow:auto;border-radius:10px}
  .src{margin-top:18px;font-size:12px;color:var(--sub);text-align:center;line-height:1.8}
  .src a{color:#93c5fd;text-decoration:none}
  @media(max-width:980px){.kpis{grid-template-columns:repeat(3,1fr)}.g1,.g2{grid-template-columns:1fr}.chart{height:300px}}
  @media(max-width:560px){.kpis{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>中美国债收益率数据面板
      <small>数据来源：东方财富数据中心 · 中国债券信息网 · 美联储</small></h1>
    <span class="badge">共 __TOTAL__ 个有效交易日 · 单位 %</span>
  </header>

  <div class="kpis">
    <div class="kpi date"><div class="lb">最新数据日期</div><div class="num">__LATEST__</div><div class="tag">数据更新至收盘</div></div>
    <div class="kpi cn"><div class="lb">中国 10 年国债收益率</div><div class="num">__CN10__<span class="un">%</span></div><div class="tag">近月均值 ~1.7</div></div>
    <div class="kpi us"><div class="lb">美国 10 年国债收益率</div><div class="num">__US10__<span class="un">%</span></div><div class="tag">美联储基准参考</div></div>
    <div class="kpi spread"><div class="lb">中美 10 年利差(中-美)</div><div class="num">__SPREAD__<span class="un">%</span></div><div class="tag">利差倒挂观察</div></div>
    <div class="kpi cn"><div class="lb">中国 2 年 / 30 年</div><div class="num">__CN2__ / __CN30__<span class="un">%</span></div><div class="tag">短端 / 长端</div></div>
    <div class="kpi us"><div class="lb">美国 2 年 / 30 年</div><div class="num">__US2__ / __US30__<span class="un">%</span></div><div class="tag">短端 / 长端</div></div>
  </div>

  <div class="grid g1">
    <div class="card">
      <h2>中美 10 年期国债收益率走势</h2>
      <div class="sub">近 __TOTAL__ 个交易日 · 同尺度对比</div>
      <div class="legend"><span><i style="background:var(--cn)"></i>中国 10 年</span><span><i style="background:var(--us)"></i>美国 10 年</span></div>
      <div id="c1" class="chart"></div>
    </div>
    <div class="card">
      <h2>中美 10 年-2 年期限利差</h2>
      <div class="sub">利差倒挂反映经济与政策预期</div>
      <div class="legend"><span><i style="background:var(--cn2)"></i>中国</span><span><i style="background:var(--us2)"></i>美国</span></div>
      <div id="c3" class="chart"></div>
    </div>
  </div>

  <div class="grid g2">
    <div class="card">
      <h2>中国国债各期限收益率</h2>
      <div class="sub">2 / 5 / 10 / 30 年期</div>
      <div id="c2cn" class="chart"></div>
    </div>
    <div class="card">
      <h2>美国国债各期限收益率</h2>
      <div class="sub">2 / 5 / 10 / 30 年期</div>
      <div id="c2us" class="chart"></div>
    </div>
  </div>

  <div class="card">
    <h2>详细数据（最近 30 个交易日）</h2>
    <div class="sub">完整数据见 CSV / Excel 表格附件</div>
    <div class="scroll"><table id="tbl"></table></div>
  </div>

  <div class="src">数据来源：<a href="https://data.eastmoney.com/cjsj/zmgzsyl.html" target="_blank">东方财富数据中心-中美国债收益率</a> · 中国国债收益率来自中国债券信息网，美国国债收益率来自美联储<br>本面板由 fetch_treasury.py 自动生成，仅供研究参考</div>
</div>

<script>
const D = __DATA__;
const base = {
  backgroundColor:'transparent', textStyle:{color:'#cbd5e1', fontFamily:'Microsoft YaHei'},
  grid:{left:46,right:18,top:30,bottom:30}, tooltip:{trigger:'axis',backgroundColor:'#0b1526',borderColor:'#334155',textStyle:{color:'#e2e8f0'},axisPointer:{type:'cross'}},
  xAxis:{type:'time',axisLine:{lineStyle:{color:'#334155'}},axisLabel:{color:'#94a3b8'},splitLine:{show:false}},
  yAxis:{type:'value',name:'收益率(%)',nameTextStyle:{color:'#94a3b8'},axisLabel:{color:'#94a3b8'},splitLine:{lineStyle:{color:'#1e293b'}}},
  dataZoom:[{type:'inside',start:0,end:100},{type:'slider',height:16,bottom:4,borderColor:'#334155',backgroundColor:'#0b1526',fillerColor:'#1e3a5f88'}]
};
const line = (data,color) => ({type:'line',showSymbol:false,smooth:true,lineStyle:{width:2,color},itemStyle:{color},data});

echarts.init(document.getElementById('c1')).setOption({
  ...base, legend:{data:['中国 10 年','美国 10 年'],textStyle:{color:'#94a3b8'},top:0},
  series:[line(D.chart1.cn10,'#38bdf8'), line(D.chart1.us10,'#fb923c')]
});

const zeroBase = {...base};
zeroBase.yAxis = {...zeroBase.yAxis, axisLabel:{formatter: v => (v>0?'+':'')+v}};
echarts.init(document.getElementById('c3')).setOption({
  ...zeroBase, legend:{data:['中国','美国'],textStyle:{color:'#94a3b8'},top:0},
  series:[{type:'line',showSymbol:false,smooth:true,lineStyle:{width:2,color:'#22d3ee'},itemStyle:{color:'#22d3ee'},areaStyle:{color:'rgba(34,211,238,.08)'},data:D.chart3.cn},
          {type:'line',showSymbol:false,smooth:true,lineStyle:{width:2,color:'#fbbf24'},itemStyle:{color:'#fbbf24'},areaStyle:{color:'rgba(251,191,36,.08)'},data:D.chart3.us}]
});

const mkTerm = (el, cols) => {
  const colors = ['#38bdf8','#22d3ee','#818cf8','#a78bfa'];
  echarts.init(document.getElementById(el)).setOption({
    ...base, legend:{type:'scroll',textStyle:{color:'#94a3b8'},top:0},
    series: cols.map((c,i)=>line(D.chart2[c], colors[i]))
  });
};
mkTerm('c2cn',['cn2','cn5','cn10','cn30']);
mkTerm('c2us',['us2','us5','us10','us30']);

const tbl = document.getElementById('tbl');
let html = '<thead><tr><th>日期</th>';
D.table_head.slice(1).forEach(h=>html+='<th>'+h.replace('国债收益率:','').replace('年增率(%)','')+'</th>');
html += '</tr></thead><tbody>';
D.table_rows.forEach(r=>{ html+='<tr>'; r.forEach((c,i)=>html+=(i===0?'<td>':'<td>')+c+'</td>'); html+='</tr>'; });
html += '</tbody>';
tbl.innerHTML = html;
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="中美国债收益率抓取+表格+图表面板")
    ap.add_argument("out_dir", nargs="?", default=".", help="输出目录（默认脚本所在目录）")
    args = ap.parse_args()

    out_dir = os.path.abspath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    print("==> 1/4 抓取数据")
    rows = fetch_all_rows()
    df = build_dataframe(rows)

    print(f"==> 2/4 保存表格（共 {len(df)} 条，{len(df.columns)} 列）")
    csv_path, xlsx_path = save_table(df, out_dir)

    print("==> 3/4 生成图表面板")
    html_path = render_dashboard(df, out_dir)

    print("==> 4/4 完成")
    print(f"  CSV  表格 : {csv_path}")
    print(f"  Excel表格 : {xlsx_path}")
    print(f"  HTML 面板 : {html_path}")


if __name__ == "__main__":
    sys.exit(main())
