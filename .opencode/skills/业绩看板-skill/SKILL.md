---
name: 业绩看板-skill
description: 蕉下直播业绩追击在线看板的搭建、数据刷新、修改、下钻交互、问题排查（基于 Excel → JSON → 深色/浅色主题 Chart.js 看板）。Use when the user asks to update the performance dashboard, regenerate dashboard_data.json from the upstream Excel, change dashboard layout or charts, fix dashboard bugs, or add drill-down interactions. Trigger keywords 业绩看板、业绩追击、dashboard、刷新看板、Top5下钻、达人下钻、机构下钻、修复看板、KPI、进度条、饼图交互、主题切换、浅色模式、深色模式、自达号、子机构、自达号下钻、自达号总业绩下探、飞书日报、personChart、focusPerson、列宽拖拽、showZidahaoDrill。
---

# 蕉下直播业绩追击看板

一个实时追踪蕉下直播团队业绩完成情况的数据看板，支持「全部」/「自达号」双数据源切换、多维度下钻分析和深色/浅色双主题切换。附带每天早上 9 点飞书群自动日报推送。

## 项目结构

```
/Users/xiaocao/CC/每日业绩自动统计/
├── dashboard.html                 # 前端看板页面（含双主题 + 图表 + 下钻）—— 本地开发用
├── public/
│   └── index.html                 # 独立部署页面（build_standalone.py 生成，内嵌数据，GH Pages 入口）
├── .gitignore                     # 排除 .DS_Store / *.pyc / .opencode/ / .playwright-mcp/
├── scripts/
│   ├── sync_dashboard.py          # 数据提取脚本（核心，读 Excel → 生成 JSON）
│   ├── watch_sync.py              # Excel 文件监控脚本（5秒轮询，自动同步+构建+部署）
│   ├── build_standalone.py        # 构建独立页面脚本（内嵌 JSON 数据，写入 public/index.html）
│   ├── daily_report.py            # 飞书日报生成 + 推送脚本
│   └── run_daily_report.sh        # launchd 包装脚本（含 webhook 环境变量）
├── data/
│   └── dashboard_data.json        # 生成的 JSON 数据文件（已提交 Git）
└── .opencode/skills/              # AI skill 配置

上游数据源：
/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx
```

## 快速使用

### 更新数据（手动）

```bash
cd "/Users/xiaocao/CC/每日业绩自动统计"
python3 scripts/sync_dashboard.py
```

### 自动同步 + 部署

`watch_sync.py` 监控 Excel 文件变化，检测到修改后自动执行三步流水线：

```
Excel 被修改 → 消抖 3 秒 → ① sync_dashboard.py（提取数据）
                              → ② build_standalone.py（生成独立页面）
                              → ③ git push（推送到 GitHub Pages）
```

watch_sync 由 crontab `@reboot` 开机启动，5 秒轮询 Excel 变化。

```bash
# 查看监控状态
ps aux | grep watch_sync | grep -v grep

# 查看同步日志
tail -f /tmp/dashboard_sync.log

# 手动启动（如进程挂了）
nohup python3 scripts/watch_sync.py > /tmp/dashboard_sync.log 2>&1 &
```

**线上看板地址**：`https://czcaizjy-lang.github.io/-/`（GitHub Pages，每次 git push 后 30 秒内生效）

### 启动 HTTP 服务器

```bash
cd "/Users/xiaocao/CC/每日业绩自动统计"
python3 -m http.server 8976
```

访问 `http://localhost:8976/dashboard.html`

### 手动推送飞书日报

```bash
FEISHU_WEBHOOK="https://open.feishu.cn/open-apis/bot/v2/hook/c44a79bd-5386-4145-ac27-50dda5a6a7fa" python3 scripts/daily_report.py
```

或直接跑包装脚本：

```bash
/bin/bash scripts/run_daily_report.sh
```

## 飞书日报系统

### 推送时间

每天早上 **9:05**（由 launchd 调度，比整点晚 5 分钟错峰）。

### 调度方式：launchd（不用 cron）

macOS 的 cron 经常因安全权限静默失败，已换成原生 `launchd`。

**配置文件**：`~/Library/LaunchAgents/com.dashboard.daily-report.plist`

**管理命令**：
```bash
# 加载
launchctl load ~/Library/LaunchAgents/com.dashboard.daily-report.plist

# 卸载
launchctl unload ~/Library/LaunchAgents/com.dashboard.daily-report.plist

# 查看状态
launchctl print gui/$(id -u)/com.dashboard.daily-report
```

**日志**：`/tmp/daily_report.log`、`/tmp/daily_report_err.log`

**Crontab 仅保留** watch_sync 的开机启动：
```
@reboot nohup /usr/bin/python3 /Users/xiaocao/CC/每日业绩自动统计/scripts/watch_sync.py > /tmp/dashboard_sync.log 2>&1 &
```

### 日报内容（5 大板块）

1. **月度总览** — 累计 GMV / 退款 / 投放 / 佣金 / 利润 / 达成率
2. **星辞昨日战报** — 仅星辞的 GMV / 支付 GMV / 退款 + 环比
3. **时间进度对标** — 时间进度 vs 达成率，应达 vs 实际 GMV
4. **达人掉量预警 Top5** — 仅主力机构（排除"其他"），环比下跌最大 5 人 + 原因推断
5. **机构昨日 GMV Top5** — 排名带奖牌

### 飞书消息格式要点

- `msg_type: "post"`（富文本）
- 每行是 `[[{"tag":"text","text":"..."}], ...]`
- ⚠️ **text 标签不支持 `style` 属性**，会报 `19002` 错误
- Webhook URL 硬编码在 `run_daily_report.sh` 和 `launchd plist` 中，不在 Python 脚本里

### daily_report.py 关键常量

```python
DAYS_IN_MONTH = 30  # 6月，跨月需修改
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
MAIN_AGENCIES = ['自达号', '集米文化', '紫语', '花开满路', '太古', '亦初', '直属']
```

## 主题系统

看板支持深色/浅色双主题，通过 CSS 变量 (`data-theme` 属性) + JS 动态切换。

### CSS 变量

| 变量 | 深色主题 | 浅色主题 | 用途 |
|------|---------|---------|------|
| `--bg` | `#0f1117` | `#f4f4f6` | 页面背景 |
| `--card` | `#1a1d27` | `#ffffff` | 卡片背景 |
| `--border` | `#2a2d3a` | `#e0e0e4` | 边框/分隔线 |
| `--text` | `#e4e4e7` | `#1c1c28` | 主文字 |
| `--text2` | `#9ca3af` | `#8e8e98` | 辅助文字 |
| `--table-head` | `#1e2130` | `#f0f0f3` | 表头背景 |
| `--accent` | `#6366f1` | `#6366f1` | 主题色 |
| `--accent2` | `#818cf8` | `#4f46e5` | 主题色辅色 |

### Chart.js 图表颜色

`getChartColors()` 函数根据 `currentTheme` 返回对应的图表配色（网格线、刻度、图例、提示框），所有图表创建时均调用此函数。

### 主题切换按钮

Header 右侧的圆形按钮：🌙（深色）/ ☀️（浅色），点击切换，选择保存在 `localStorage`（key: `dashboardTheme`）。

## 数据提取逻辑

### Excel 结构

上游 Excel 包含多个 Sheet：

- **星辞业绩**: 主数据表，包含达人汇总（col 1-25）、机构汇总（col 27-39）、汇总行（row 2）
- **星辞自达号业绩**: 自达号专用数据表（132 行），col 10 为子机构分类（花开自达号/集米自达号/太古自达号/九三自达号/直属自达号），col 13 直播总GMV，col 20 直播结算GMV，col 25 消耗金额，col 26 ROI，cols 31-47 子机构汇总
- **总自达号业绩**: 自达号逐日明细表（4938 行），col 3 机构归属，含每日 GMV/退款/结算/消耗/ROI
- **6月直播数据**: 逐日逐条直播明细，col 3 是抖音号、col 4 是日期、col 26 GMV、col 27 支付金额、col 32 退款、col 44 投放消耗(店铺绑定)、col 45 投放消耗(店铺被投)
- **久酒业绩 / 雅宁业绩 / 星辞业绩**: 个人 Sheet，按抖音号匹配到 6月直播数据，统计三人分天业绩（支付、GMV、退款）

### 关键列对应关系

| 数据类型 | Sheet | 抖音号列 | 日期列 | 数值列 |
|---------|-------|---------|--------|--------|
| 达人汇总 | 星辞业绩 | col 2 | - | col 1-25 |
| 机构汇总 | 星辞业绩 | - | - | col 27-39 |
| 自达号达人 | 星辞自达号业绩 | col 2 | - | col 10 (子机构), 13 (GMV), 20 (结算), 25 (消耗), 26 (ROI) |
| 自达号子机构 | 星辞自达号业绩 | - | - | col 31-47 (rows 3-7) |
| 直播明细 | 6月直播数据 | col 3 | col 4 | col 26 (GMV), 27 (支付), 32 (退款), 44 (投放绑定), 45 (投放被投) |
| 三人分天 | 久酒/雅宁/星辞业绩 | col 2 | - | 匹配到 6月直播数据汇总 |

**关键口径**：

- **每日消耗** = `max(col 44 投放消耗店铺绑定, col 45 投放消耗店铺被投)` —— 取两者中较大值，避免漏算
- **达人月度消耗** = Σ(该达人所有日消耗)，从 `6月直播数据` 按日汇总，非从 Excel 汇总表直接读取
- **机构/子机构月度消耗** = Σ(该机构所有达人月度消耗)
- **总月度消耗** = Σ(所有达人月度消耗)
- **ROI** = GMV / 消耗（月度 OR 分日）
- **分日子机构 ROI** = Σ(子机构当日所有达人 GMV) / Σ(子机构当日所有达人消耗)，加权计算，非简单平均

所有消耗和 ROI 均由 `sync_dashboard.py` 在 section 4b 中统一重算，不再依赖 Excel 汇总表中的预计算值。col 34 `预估佣金支出` 是另一口径，不要混用。

### JSON 数据结构

生成的 `dashboard_data.json` 包含：

```javascript
{
  summary: {},      // 汇总数据（GMV、退款、结算率等）
  anchors: [],      // 达人列表（含直播GMV、佣金率等）
  agencies: [],     // 机构列表
  zidahao: {},      // 自达号专用数据（summary/sub_agencies/anchors/daily_*/top5_by_sub）
  trend: {
    dates: [],                    // 日期数组 ["06/01", "06/02", ...]
    daily_total_gmv: [],          // 每日总GMV
    daily_total_paid: [],         // 每日总支付金额
    daily_total_refund: [],       // 每日总退款
    agencies: [],                 // 机构名称数组
    daily_by_agency: {},          // {机构名: [每日GMV]}
    agency_totals: [],            // 各机构总GMV
    pie_data: [],                 // 机构饼图数据
    person_daily: {               // 久酒/雅宁/星辞 分天支付
      dates: [],
      久酒: [], 雅宁: [], 星辞: []
    },
    person_daily_gmv: {           // 久酒/雅宁/星辞 分天 GMV
      dates: [],
      久酒: [], 雅宁: [], 星辞: []
    },
    person_daily_refund: {        // 久酒/雅宁/星辞 分天退款
      dates: [],
      久酒: [], 雅宁: [], 星辞: []
    },
    anchor_daily_paid: {          // 达人分天支付（点击下钻用）
      "抖音号": {"06/01": 123.45, "06/02": 234.56}
    },
    anchor_daily_roi: {           // 达人分天 ROI (GMV / 投放消耗，每日消耗用 max(col44,col45))
      "抖音号": {"06/01": 8.92, "06/02": 11.52}
    },
    agency_top5_anchors: {        // 机构 Top5 达人分天数据
      "自达号": [
        {name: "达人A", douyin_id: "xxx", daily_paid: [12.3, 23.4, ...]},
        ...  // 共5条
      ]
    },
    person_anchor_detail: {       // 人员下探：每人下属所有达人分天支付明细
      "久酒": [
        {name: "达人A", douyin_id: "xxx", daily_paid: [1.2, 3.4, ...]},
        ...  // 63人（按支付额降序）
      ],
      "雅宁": [...],   // 197人
      "星辞": [...]    // 262人
    }
  },
  zidahao: {                     // 自达号专用数据
    summary: {                   // 自达号整体汇总
      直播GMV: 5203736.68,
      直播结算GMV: 2504575.19,
      消耗金额: 520469.04,
      ROI: 9.998
    },
    sub_agencies: [],            // 5 个子机构汇总（花开/集米/太古/九三/直属自达号）
    anchors: [],                 // 130 个自达号达人（机构 = 子机构名）
    daily_gmv: [],               // 自达号每日 GMV
    daily_paid: [],              // 自达号每日支付
    daily_refund: [],            // 自达号每日退款
    daily_by_sub: {              // 子机构每日 GMV
      "花开自达号": [], "集米自达号": [], ...
    },
    pie_data: [],                // 子机构饼图数据
    top5_by_sub: {               // 子机构 Top5 自达号达人分天数据
      "花开自达号": [
        {name: "达人A", douyin_id: "xxx", daily_paid: [12.3, ...]},
        ...
      ]
    },
    daily_roi_by_sub: {          // 子机构每日加权 ROI（ΣGMV/Σ消耗，非简单平均）
      "花开自达号": [10.5, 9.8, null, ...],
      "集米自达号": [11.2, 10.3, null, ...],
      ...
    },
    anchor_detail: [             // 自达号下探：所有自达号达人分天支付明细（按支付额降序）
      {name: "达人A", douyin_id: "xxx", daily_paid: [12.3, 23.4, ...]},
      ...  // 137人
    ]
  }
}
```

**关键**：`agency_top5_anchors` 是全部视图的字段名；`zidahao.top5_by_sub` 是自达号视图的字段名。

## 前端功能

### 数据源切换

导航栏左侧新增「**全部** / **自达号**」切换按钮（橙色选中态），点击切换整个看板的数据源：

- **全部模式**（默认）：全局数据，3 张 KPI 卡片（GMV/结算/退款）、时间进度条、4 个图表、256 达人
- **自达号模式**：自达号专属数据，4 张 KPI 卡片（GMV/结算/消耗/ROI —— ROI 卡片可点击下钻）、**无时间进度**、3 个图表（总趋势/子机构趋势/子机构饼图）、137 自达号达人

切换逻辑由 `switchDataSource(source)` 统一管理，设置 `currentDataSource` 全局状态后销毁所有图表并重新渲染。

### 图表（全部模式）

- **每日总业绩趋势**（折线图）: 总GMV、支付金额、退款金额
- **每日机构业绩趋势**（折线图）: 7个主力机构，支持点击线/图例聚焦单个机构，再点恢复全部
- **久酒/雅宁/星辞分天业绩**（折线图）: 三人对比；标题右侧有 4 个按钮（全部/久酒/雅宁/星辞），点击聚焦单人，再点同一按钮恢复全部；**点击数据点可弹出该人员下所有达人的当日 vs 前日变化明细（人员下探）**，支持全部/仅下降/仅上涨筛选
- **机构业绩占比**（饼图/环形图）: 点击扇区可弹出该机构 Top5 达人弹窗

### 图表（自达号模式）

- **自达号总业绩趋势**（折线图）: 自达号总 GMV/支付/退款，标签含「自达号」后缀；**点击数据点可弹出所有自达号达人的当日 vs 前日变化明细（`showZidahaoDrill()`）**，支持全部/仅下降/仅上涨筛选
- **自达号子机构趋势**（折线图）: 5 条子机构线（花开/集米/太古/九三/直属自达号），支持点击聚焦
- **子机构业绩占比**（饼图/环形图）: 5 个子机构扇区，点击弹出该子机构 Top5 自达号达人分天趋势弹窗
- **久酒/雅宁/星辞图表**：隐藏（`#personTrendCard` display:none）

子机构图表由 `renderZidahaoCharts()` 函数独立渲染，数据结构来自 `DATA.zidahao`。

**自达号 ROI KPI 下钻**：点击 ROI KPI 卡片（带「点击看趋势」提示），弹出各子机构每日 ROI 趋势折线图（5 条线，`showZidahaoRoiDrill()`）。数据来自 `zidahao.daily_roi_by_sub`，口径为 ΣGMV/Σ消耗 加权计算，支持点击聚焦/取消聚焦。

### 交互

- **数据源切换**: 点击「全部」/「自达号」按钮，销毁全部图表后重新渲染（`switchDataSource()`）
- **主题切换**: Header 右侧 🌙/☀️ 按钮，深色/浅色双主题，自动保存到 localStorage
- **机构趋势图**: 点击某条线只显示该机构，再点恢复全部（`window.__agencyFocusMode`）
- **三人业绩按钮**（`focusPerson()`）: 图表标题右侧 4 个 `.person-btn`，点击聚焦单人曲线，再点同一按钮恢复全部。状态保存在 `window.__personFocused`，按钮样式跟随切换
- **人员下探**（`showPersonDrill(personName, dateIndex)`）: 点击三人图表任意数据点，弹出该人员下所有达人的当日 vs 前日变化明细表。按掉量降序（最大掉量排第一），红色 ↓ 标注下降、绿色 ↑ 标注上涨。顶部可切换「全部 / ↓ 仅下降 / ↑ 仅上涨」筛选（`filterPersonDrill()`）。数据来自 `trend.person_anchor_detail`，达人昵称优先从人员 sheet A 列读取
- **自达号总业绩下探**（`showZidahaoDrill(dateIndex)`）: 点击自达号总业绩趋势图任意数据点，弹出所有自达号达人（137人）的当日 vs 前日变化明细表。复用人员下探弹窗和筛选按钮逻辑，按掉量降序排列。数据来自 `zidahao.anchor_detail`
- **饼图下钻**: 点击机构扇区弹出 Top5 达人分天曲线（5 种颜色，支持弹窗内点击聚焦）。`showAgencyDrill()` 根据 `currentDataSource` 自动读取 `agency_top5_anchors`（全部）或 `top5_by_sub`（自达号）
- **达人下钻**: 点击达人列表中的「直播GMV」单元格弹出该达人每日支付曲线
- **ROI 下钻**（2 种入口）：
  - 达人表 ROI 列：点击弹出该达人每日 ROI 趋势（GMV/投放，琥珀色线条）—— `showRoiDrill()`
  - 自达号 ROI KPI 卡片：点击弹出各子机构每日 ROI 趋势（5 条线，加权口径）—— `showZidahaoRoiDrill()`
- **列拖拽排序**: 达人表头支持 HTML5 拖拽，拖动列头即可调整列顺序（`DISPLAY_COLS` 运行时可变）
- **列宽拖拽**: 鼠标移到表头右边缘，光标变 `col-resize`，拖拽调整列宽，宽度保存到 localStorage
- **搜索**: 输入框回车或点「查询」按钮，不实时过滤
- **备注编辑**: 达人名下方备注列可编辑，保存到 localStorage
- **GMV 目标**: 可自定义目标 GMV（万），进度条根据完成情况变色（绿/橙/红），保存在 `localStorage.targetGmvWan`（自达号模式下不显示）

### 关键 JS 全局变量

| 变量 | 用途 |
|------|------|
| `currentDataSource` | 当前数据源：`'all'`（全部）或 `'zidahao'`（自达号） |
| `ZDH_SUB_AGENCIES` | 自达号 5 子机构数组：花开/集米/太古/九三/直属自达号 |
| `ZDH_PIE_COLORS` | 自达号饼图配色：`['#818cf8','#22c55e','#f59e0b','#ec4899','#14b8a6']` |
| `window.__totalChart` | 总业绩趋势图表实例 |
| `window.__agencyChart` / `window.__agencyFocusMode` | 机构趋势图表 / 聚焦状态 |
| `window.__personChart` | 三人分天图表实例（自达号模式下为 null） |
| `window.__personFocused` | 当前聚焦的人名（`'all'` / `'久酒'` / `'雅宁'` / `'星辞'`） |
| `window.__agencyPieChart` | 机构饼图实例 |
| `window.__drillChart` | 达人下钻弹窗图表实例（GMV / ROI 共用） |
| `window.__agencyDrillChart` / `window.__agencyDrillFocusMode` | 机构 Top5 / 自达号 ROI 趋势弹窗图表 / 聚焦状态 |
| `MAIN_AGENCIES` | 主力 7 机构：自达号、集米文化、紫语、花开满路、太古、亦初、直属 |
| `showZidahaoRoiDrill()` | 自达号 ROI KPI 下钻函数，弹出子机构每日 ROI 趋势（`daily_roi_by_sub`） |
| `showPersonDrill(personName, dateIndex)` | 人员下探函数，弹出该人员所有达人的当日 vs 前日变化明细表 |
| `filterPersonDrill(filter)` | 人员下探筛选按钮（'all' / 'down' / 'up'），过滤表格并调整排序 |
| `closePersonDrillModal()` | 关闭人员下探弹窗 |
| `showZidahaoDrill(dateIndex)` | 自达号总业绩下探函数，弹出所有自达号达人的当日 vs 前日变化明细表 |
| `window.__personDrillRows` / `__personDrillPrevDate` / `__personDrillClickedDate` | 人员下探 / 自达号下探缓存的原始数据（共用） |

### 达人表显示规则

- `DISPLAY_COLS` 是运行时可变的数组（支持拖拽调整）
- **ROI 列**：动态计算 `直播GMV / 投放消耗金额`，显示为数字（如 `8.9`），无投放时显示 `-`。**不显示颜色方块**，字体 15px 加粗
- **目前结算率**：保留颜色方块（绿 ≥40% / 橙 20-40% / 红 <20%），字体 15px 加粗
- **居中对齐**：从「开播天数」列起的所有列（表头 + 数据单元格）均 `text-align: center`
- **抖音号处理**：列拖拽和 ROI 下钻时，必须用 `String(douyinId).replace(/'/g, "\\'")` 转义，因为抖音号可能是数字类型

## 部署上线

### 线上地址

| 平台 | 地址 | 说明 |
|------|------|------|
| GitHub Pages | `https://czcaizjy-lang.github.io/-/` | 主站，每次 git push 自动更新 |
| 本地预览 | `http://localhost:8976/dashboard.html` | 开发调试用 |

### 自动部署（推荐）

`watch_sync.py` 内置自动部署：Excel 变化 → 数据同步 → 构建页面 → `git push`。只要电脑开机 + watch_sync 在跑，线上看板就自动保持最新。

```
监控 Excel 变化
    │
    └── 变化检测到
          ├── ① python3 scripts/sync_dashboard.py    → 生成 data/dashboard_data.json
          ├── ② python3 scripts/build_standalone.py  → 生成 public/index.html
          └── ③ git add + git commit + git push       → 推送到 GitHub Pages
```

### 手动部署

```bash
cd "/Users/xiaocao/CC/每日业绩自动统计"

# 第 1 步：提取数据
python3 scripts/sync_dashboard.py

# 第 2 步：构建独立页面（内嵌 JSON，可直接部署）
python3 scripts/build_standalone.py

# 第 3 步：推送到 GitHub Pages
git add data/dashboard_data.json public/index.html
git commit -m "📊 手动数据更新 $(date +'%Y-%m-%d %H:%M')"
git push origin main
```

### build_standalone.py 工作原理

`dashboard.html` 是模板（开发版），通过 `fetch` 异步加载 `data/dashboard_data.json`。`build_standalone.py` 将 JSON 数据直接内嵌到 HTML 中，生成完全独立的静态页面：

1. 读取 `dashboard.html` 模板
2. 找到 `fetch('data/dashboard_data.json?...')` 语句
3. 替换为 `<script id="inline-data" type="application/json">{数据}</script>` 内嵌块
4. 写入 `public/index.html`（GitHub Pages 入口）

这样生成的页面无需本地服务器，可直接通过 `file://` 打开或部署到任何静态托管平台。

### 部署架构

```
本地 Mac                            GitHub (czcaizjy-lang/-)
┌─────────────┐                    ┌──────────────────────┐
│ Excel (.xlsx)│                   │  main 分支            │
│      ↓       │                   │  ├── public/          │
│ sync_data.py │                   │  │   └── index.html   │
│      ↓       │   git push        │  └── data/            │
│ build_stand- │ ───────────────→  │      └── dashboard_   │
│ alone.py     │                   │         data.json     │
│      ↓       │                   └──────────┬───────────┘
│ watch_sync.py│                              │
└─────────────┘                      GitHub Pages (public/)
                                      ┌──────────────────┐
                                      │ czcaizjy-lang.   │
                                      │ github.io/-/     │
                                      └──────────────────┘
```

## 常见问题

### 1. 数据加载失败

前端 `window.DATA` 为 `undefined`。

**排查**:
```javascript
// 浏览器控制台
fetch('dashboard_data.json?t=' + Date.now()).then(r => r.json()).then(d => {
  window.DATA = d;
  render();
}).catch(e => console.error(e));
```

### 2. JSON 生成错误

脚本报 `KeyError: 'Worksheet xxx does not exist.'`

**解决**: 检查上游 Excel 文件的 Sheet 名称是否与脚本中硬编码的一致：
- `星辞业绩`
- `6月直播数据`
- `久酒业绩` / `雅宁业绩` / `星辞业绩`

### 3. 日期类型错误

脚本报 `unsupported operand type(s)` 与日期相关的类型错误。

**原因**: Excel 中日期列可能是字符串 `"2026/06/14 07:01:38"` 而非 datetime 对象。

**解决**: 使用字符串解析：
```python
date_key = str(dt_val)[:10].replace('/', '-')  # 提取 "2026-06-14"
```

### 4. 抖音号类型

脚本报 `TypeError: unsupported operand type(s)` 与抖音号相关的类型错误。

**原因**: 抖音号可能是字符串 `"xxx"` 或数字 `123456`。

**解决**: 始终用字符串比较：
```python
douyin_id_str = str(douyin_id)
```

### 5. 字段名错误

前端报 `Cannot read properties of null` 或 `undefined`。

**排查**: 用浏览器开发工具检查 JSON 字段名是否正确：
```javascript
// 浏览器控制台
console.log(DATA.trend.agency_top5_anchors);  // 机构 Top5
console.log(DATA.trend.person_daily);          // 三人分天支付
console.log(DATA.trend.person_daily_gmv);      // 三人分天 GMV
console.log(DATA.trend.person_daily_refund);   // 三人分天退款
console.log(DATA.trend.person_anchor_detail);  // 人员下探达人明细
console.log(DATA.trend.anchor_daily_paid);     // 达人分天
```

### 6. 自动同步不工作

**排查步骤**:
```bash
# 检查监控进程是否运行
ps aux | grep watch_sync | grep -v grep

# 查看日志
tail -20 /tmp/dashboard_sync.log

# 手动测试同步
cd "/Users/xiaocao/CC/每日业绩自动统计"
python3 scripts/sync_dashboard.py

# 如进程挂了，手动重启
nohup python3 scripts/watch_sync.py > /tmp/dashboard_sync.log 2>&1 &
```

### 7. 主题切换后图表颜色没变

主题切换时 `applyTheme()` 会调用 `renderTrendCharts()` 重绘所有图表。如果弹窗（下钻图表）处于打开状态，关闭后重新打开即可获得新主题颜色。

### 8. 飞书日报没推送

**可能原因**：

- **cron 静默失败**（macOS 常见）：已切换为 launchd，检查 `launchctl print gui/$(id -u)/com.dashboard.daily-report`
- **launchd job 未加载**：`launchctl load ~/Library/LaunchAgents/com.dashboard.daily-report.plist`
- **脚本报错**：查看 `/tmp/daily_report.log` 和 `/tmp/daily_report_err.log`
- **webhook 失效**：手动跑一次 `daily_report.py` 确认

### 9. 三个人分天业绩按钮不生效

- 确认 `focusPerson()` 函数存在（搜索 dashboard.html）
- 确认 `.person-btn` 按钮存在
- 打开浏览器控制台，检查 `window.__personChart` 是否已创建
- 强制刷新页面 (`Cmd+Shift+R`)，清除缓存

### 10. 浏览器缓存看到旧版本

- 按 `Cmd+Shift+R` 强制刷新
- 或访问 `http://localhost:8976/dashboard.html?v=2` 绕过缓存

### 11. 自达号模式切换后图表空白 / Canvas 报错

**原因**：Chart.js 不允许在同一 canvas 上创建第二个实例，必须先 `.destroy()`。

**解决**：`switchDataSource()` 已内置销毁逻辑，正常情况下不应出现。如出现，刷新页面即可。排查时检查 `currentDataSource` 变量值。

### 12. 自达号数据不显示或为 0

**可能原因**：
- `sync_dashboard.py` 未正确读取 `星辞自达号业绩` sheet
- JSON 中缺少 `zidahao` key

**排查**：
```javascript
// 浏览器控制台
console.log(DATA.zidahao);             // 检查 zidahao 数据是否存在
console.log(DATA.zidahao.summary);     // 汇总数据
console.log(DATA.zidahao.anchors.length);  // 应为 130
```

### 13. 自达号饼图点击没有弹出 Top5 达人

**确认**：`showAgencyDrill()` 在自达号模式下从 `DATA.zidahao.top5_by_sub` 读取数据。如果某个子机构没有直播数据，会提示"暂无 Top 5 达人分天数据"。

### 14. 跨月数据更新

上游 Excel 路径中包含月份（`6月业绩`），跨月后需要更新：
- `sync_dashboard.py` 中的 `XLSX_PATH`
- `watch_sync.py` 中的 `XLSX_PATH`
- `daily_report.py` 中的 `DAYS_IN_MONTH`
- 看板标题中的月份（`dashboard.html` h1）
- 建议每月初检查并更新路径

## Git 托管

- **仓库**：`git@github.com:czcaizjy-lang/-.git`
- **线上看板**：`https://czcaizjy-lang.github.io/-/`（GitHub Pages）
- **分支**：main
- **SSH Key**：`~/.ssh/id_ed25519`（czcaizjy@gmail.com）
- **Git 配置**：user.name = xiaocao, user.email = czcaizjy@gmail.com
