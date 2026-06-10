# 股指期货吃贴水与跨期套利策略研究

本项目研究中证 500 股指期货（IC）和中证 1000 股指期货（IM）的吃贴水策略、跨期套利策略及二者的组合效果。研究从原始 Excel 数据出发，在 Jupyter Notebook 中完成数据清洗、因子构造、回测、统计检验、收益归因和图表输出，并形成 Word 研究报告。

报告采用偏券商金工研报的口径：结论不过度包装，重点区分“样本内表现”“统计检验结果”和“经济解释”。吃贴水策略在样本内相对指数有明显收益增厚，但本质仍是带有权益方向暴露的期货多头替代，并不能简单理解为市场中性套利。

## 核心结论

| 项目 | IC | IM |
|---|---:|---:|
| Always 年化收益 | 9.60% | 15.78% |
| 指数年化收益 | 0.83% | 5.31% |
| Always Sharpe | 0.23 | 0.42 |
| Always 最大回撤 | -59.39% | -46.31% |
| Bootstrap SR 差异 95%CI | [0.0452, 0.5907] | [0.1284, 0.4731] |
| Bootstrap p 值 | 0.023 | <0.001 |
| Always β | 0.9616 | 1.2387 |
| 残差年化 | 9.62% | 9.38% |

主要判断：

- 吃贴水策略相对指数有收益增厚，来源主要是贴水结构下的基差收敛与展期补偿。
- 策略回撤较深，IC Always 最大回撤接近 -60%，不能定位为低波动或绝对收益策略。
- 无截距归因显示 Always 策略 β 接近或高于 1，说明方向暴露是收益解释中不可忽略的部分。
- 跨期套利在当前日频 T+1 框架下表现较弱，64 组参数组合均未取得正 Sharpe；加杠杆不会改变收益来源，只会同步放大亏损和回撤。
- IC 与 IM 吃贴水策略日收益高度相关，品种分散化价值有限，组合层面的改进更依赖引入低相关策略。

## 项目结构

```text
data/
  000852.xlsx
  000905.xlsx
  ic_data.xlsx
  im_data.xlsx

notebooks/
  股指期货套利策略研究.ipynb

output/
  figures/
    fig_roll_yield.png
    fig_nav.png
    fig_annual.png
    fig_drawdown.png
    fig_heatmap.png
    fig_calendar_leverage.png
    fig_correlation.png
    fig_bootstrap.png
    fig_attribution.png
    fig_cost_stress.png
  tables/
    strategy_summary.csv
    combination_results.csv
    spread_scan.csv
    spread_leverage.csv
    bootstrap_validation.csv
    attribution_summary.csv
  validation_summary.csv

report/
  report.md
  build_report_docx.py
  assets/
  股指期货套利策略研究报告.docx
```

## 复现方式

在项目根目录执行：

```bash
python -m jupyter nbconvert --to notebook --execute --inplace "notebooks/股指期货套利策略研究.ipynb" --ExecutePreprocessor.timeout=1200
```

本地环境中也可以使用：

```bash
D:\Anaconda\envs\QuantEnv\python.exe -m jupyter nbconvert --to notebook --execute --inplace "notebooks/股指期货套利策略研究.ipynb" --ExecutePreprocessor.timeout=1200
```

执行完成后会更新 `output/figures/`、`output/tables/` 和 `output/validation_summary.csv`。当前项目不依赖 `.pkl` 缓存文件。

## 主要图表

- `fig_roll_yield.png`：主力合约持有段收益分布
- `fig_nav.png`：策略净值曲线
- `fig_annual.png`：年度收益拆分
- `fig_drawdown.png`：回撤路径
- `fig_heatmap.png`：参数敏感性热力图
- `fig_calendar_leverage.png`：跨期套利杠杆压力测试
- `fig_correlation.png`：IC-IM 日收益相关性
- `fig_bootstrap.png`：Bootstrap 显著性检验
- `fig_attribution.png`：收益归因分解
- `fig_cost_stress.png`：交易成本压力测试

## 交付文件

- 研究报告：`report/股指期货套利策略研究报告.docx`
- Markdown 版本：`report/report.md`
- 研究过程：`notebooks/股指期货套利策略研究.ipynb`
- 结果表：`output/tables/`
- 复现检查：`output/validation_summary.csv`

## 风险提示

本项目基于历史样本回测，不构成投资建议。股指期货策略受权益市场方向、保证金制度、交易限制、分红影响、贴水结构变化和样本区间选择影响较大。IM 样本自 2022 年开始，覆盖周期较短，相关结论应低于 IC 样本的置信度。
