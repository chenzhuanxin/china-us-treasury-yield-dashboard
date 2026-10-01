# 中美国债收益率数据面板 (China-US Treasury Yield Dashboard)

从东方财富数据中心抓取中美国债收益率数据，输出表格并生成 ECharts 图表面板。

## 文件说明

| 文件 | 说明 |
|------|------|
| `fetch_treasury.py` | 数据抓取脚本：抓取东方财富接口全量数据，输出 CSV / Excel 表格，并生成 HTML 数据面板 |
| `中美国债收益率_数据面板.html` | 数据图表面板（ECharts），含 KPI 卡片、走势图、期限利差、明细表 |
| `index.html` | 跳转页，自动转到数据面板 |

## 数据来源

- 页面：<https://data.eastmoney.com/cjsj/zmgzsyl.html>
- 接口：`RPTA_WEB_TREASURYYIELD`（东方财富数据中心）
- 中国国债收益率来自中国债券信息网，美国国债收益率来自美联储

## 使用

```bash
pip install requests pandas openpyxl
python fetch_treasury.py            # 默认输出到脚本所在目录
python fetch_treasury.py D:/out     # 指定输出目录
```

生成产物：`中美国债收益率.csv`、`中美国债收益率.xlsx`、`中美国债收益率_数据面板.html`

## 面板访问

- GitHub Pages：<https://chenzhuanxin.github.io/china-us-treasury-yield-dashboard/>

> 仅供研究参考，不构成任何投资建议。
