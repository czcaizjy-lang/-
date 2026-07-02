---
name: 业绩看板-skill
description: 蕉下直播业绩追击在线看板的搭建、数据刷新、修改、下钻交互、问题排查（基于 Excel → JSON → 深色/浅色主题 Chart.js 看板）。Use when the user asks to update the performance dashboard, regenerate dashboard_data.json from the upstream Excel, change dashboard layout or charts, fix dashboard bugs, or add drill-down interactions. Trigger keywords 业绩看板、业绩追击、dashboard、刷新看板、Top5下钻、达人下钻、机构下钻、修复看板、KPI、进度条、饼图交互、主题切换、浅色模式、深色模式、自达号、子机构、自达号下钻、自达号总业绩下探、飞书日报、personChart、focusPerson、列宽拖拽、showZidahaoDrill、应达GMV、showExpectedGmv、时均GMV、bar-cell、showOverallDrill、全部达人下探、整体趋势下探、showPersonDrill、showOtherDrill、其他下探、all_anchor_daily、person_anchor_detail、消耗口径、被投优先、live_name_map、agencyDrillShowAll、updateShowAllButton、划掉交互、隐藏趋势线、7月、月度过滤、daily_by_agency。
---

# 蕉下直播业绩追击看板

一个实时追踪蕉下直播团队业绩完成情况的数据看板，支持「全部」/「自达号」双数据源切换、多维度下钻分析和深色/浅色双主题切换。

## 项目结构

```
/Users/xiaocao/CC/每日业绩自动统计/
├── dashboard.html                 # 前端看板页面（含双主题 + 图表 + 下钻）—— 本地开发用
├── public/
│   └── index.html                 # 独立部署页面（build_standalone.py 生成，内嵌数据，GH Pages 入口）
├── scripts/
│   ├── sync_dashboard.py          # 数据提取脚本（核心，读 Excel → 生成 JSON）
│   ├── auto_sync.sh               # 自动同步脚本（launchd WatchPaths 触发，含消抖）
│   ├── build_standalone.py        # 构建独立页面脚本（内嵌 JSON 数据，写入 public/index.html）
│   ├── watch_sync.py              # （已弃用）旧轮询监控脚本，已被 launchd WatchPaths 替代
│   ├── daily_report.py            # （已弃用）飞书日报脚本，已于 2026-06-24 停用
│   └── run_daily_report.sh        # （已弃用）飞书日报 launchd 包装脚本
├── data/
│   ├── dashboard_data.json        # 生成的 JSON 数据文件（1.5 MB）
│   ├── sync.log                   # auto_sync.sh 同步日志
│   └── server.log                 # 本地 HTTP 服务器日志
└── .opencode/skills/              # AI skill 配置

上游数据源:
/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx
```

## 快速使用

### 更新数据（手动）

```bash
cd "/Users/xiaocao/CC/每日业绩自动统计"
python3 scripts/sync_dashboard.py
```

此命令会自动执行数据提取 + 构建页面两步。

### 自动同步 + 部署（launchd WatchPaths）

采用 macOS 原生内核级文件监控，**Excel 保存时自动触发，其他时间零进程、零 CPU、零内存**。

```
Excel 保存 → macOS 内核 WatchPaths 检测到变化
         → launchd 启动 auto_sync.sh
              → 消抖 5~10 秒（双重防抖）
              → ① sync_dashboard.py（提取数据，通过 osascript 获取桌面文件权限）
              → ② build_standalone.py（生成独立页面）
              → ③ git push（推送到 GitHub Pages）
         → 脚本退出，恢复零占用
```

### launchd 任务管理

| 任务名 | 作用 | 触发方式 |
|--------|------|---------|
| `com.dashboard.sync` | Excel 保存 → 同步 + 构建 + 推送 | WatchPaths 文件变化 |
| `com.dashboard.local-server` | 本地 HTTP 服务器 :8976 | 开机自启，常驻（~10MB） |

```bash
# 查看任务状态
launchctl list | grep dashboard

# 手动触发一次同步（测试用）
bash /Users/xiaocao/CC/每日业绩自动统计/scripts/auto_sync.sh

# 查看同步日志
tail -f /Users/xiaocao/CC/每日业绩自动统计/data/sync.log

# 卸载/重载
launchctl unload ~/Library/LaunchAgents/com.dashboard.sync.plist
launchctl load ~/Library/LaunchAgents/com.dashboard.sync.plist
```

### 本地看板

本地 HTTP 服务器由 launchd 自动管理（`com.dashboard.local-server`），端口 8976，开机自启。

访问 `http://localhost:8976/dashboard.html`

**线上看板地址**：`https://czcaizjy-lang.github.io/-/`（GitHub Pages，每次 git push 后 30 秒内生效）

### macOS 权限要求

launchd 启动的进程默认无权访问桌面文件。**必须将 `/bin/bash` 加入「系统设置 → 隐私与安全性 → 完全磁盘访问权限」**，否则 sync_dashboard.py 会报 `PermissionError`。

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

- **星辞业绩**: **花名册**（仅读 col 1 昵称、col 2 抖音号、col 9 机构），不再读取月度汇总数值列
- **星辞自达号业绩**: **自达号花名册**（仅读 col 1 昵称、col 2 抖音号、col 10 子机构），不再读取月度汇总数值列
- **X月直播数据**: 逐日逐条直播明细（自动识别名称含「直播数据」的 Sheet），col 3 是抖音号、col 4 是日期、col 6 直播时长、col 26 GMV、col 27 支付金额、col 32 退款、col 34 佣金支出、col 44 投放消耗(店铺绑定)、col 45 投放消耗(店铺被投)
- **久酒业绩 / 雅宁业绩**: 个人 Sheet，仅读 col 1（昵称）+ col 2（抖音号）作为花名册，按抖音号匹配到日流水统计三人分天业绩（支付、GMV、退款）

### 关键列对应关系

| 数据类型 | Sheet | 抖音号列 | 日期列 | 数值列 |
|---------|-------|---------|--------|--------|
| 达人汇总 | 星辞业绩 | col 2 | - | col 1-25 |
| 机构汇总 | 星辞业绩 | - | - | col 27-39 |
| 自达号达人 | 星辞自达号业绩 | col 2 | - | col 10 (子机构), 13 (GMV), 20 (结算), 25 (消耗), 26 (ROI) |
| 自达号子机构 | 星辞自达号业绩 | - | - | col 31-47 (rows 3-7) |
| 直播明细 | X月直播数据 | col 3 | col 4 | col 26 (GMV), 27 (支付), 32 (退款), 44 (投放绑定), 45 (投放被投) |
| 三人分天 | 久酒/雅宁/星辞业绩 | col 2 | - | 匹配到 X月直播数据汇总 |

**关键口径**：

- **月份自动识别**：从日流水最大日期提取月份（如 `2026-06`），所有月度聚合自动按该月过滤
- **所有数值均从日流水实时计算**：GMV、支付、退款、消耗、佣金、开播天数、日均时长等均从「X月直播数据」聚合，不再读取 Excel 汇总表中的预计算值
- **结算GMV** = 支付金额 - 退款GMV
- **每日消耗** = **优先取 col 45（投放消耗店铺被投），为 0 时回退到 col 44（投放消耗店铺绑定）**
- **安全过滤**：GMV=0 的场次，消耗不计入日汇总（`ad_cost if gmv > 0 else 0`）
- **ROI** = GMV / 消耗（月度 OR 分日）
- **趋势窗口** = 近 30 天滚动（以日流水最新日期为终点前推 30 天）
- **分日子机构 ROI** = Σ(子机构当日所有达人 GMV) / Σ(子机构当日所有达人消耗)，加权计算，非简单平均
- **「其他」= 全部日流水达人 - 久酒 - 雅宁 - 星辞**，在三人分天图中作为第 4 条灰色虚线展示
- **达人昵称优先级**：日流水 B 列 > 花名册 A 列，日流水覆盖最全
- 已移除字段：团长服务费、预测税后收入、预测直播结算GMV、预估利润、预估6月收入、当前收入达成率、4月结算率、达人备注

### JSON 数据结构

生成的 `dashboard_data.json` 包含：

```javascript
{
  summary: {},      // 汇总数据（当月聚合：GMV、退款、结算、佣金、消耗）
  anchors: [],      // 达人列表（当月聚合：GMV/退款/结算/消耗/佣金/开播天数/时长/ROI/结算率）
  agencies: [],     // 机构列表（当月聚合）
  zidahao: {},      // 自达号专用数据（summary/sub_agencies/anchors/daily_*/top5_by_sub）
  trend: {
    dates: [],                    // 近30天日期数组 ["05/25", "05/26", ...]
    daily_total_gmv: [],          // 每日总GMV（万）
    daily_total_paid: [],         // 每日总支付金额（万）
    daily_total_refund: [],       // 每日总退款（万）
    agencies: [],                 // 机构名称数组
    daily_by_agency: {},          // {机构名: [每日GMV（万）]}
    agency_totals: number,        // 各机构总GMV汇总（万）
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
      "抖音号": {"05/25": 123.45, "05/26": 234.56}
    },
    anchor_daily_roi: {           // 达人分天 ROI
      "抖音号": {"05/25": 8.92, "05/26": 11.52}
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
        ...
      ],
      "雅宁": [...],
      "星辞": [...]
    },
    all_anchor_daily: [           // 全部达人下探：所有达人分天支付明细
      {name: "达人A", douyin_id: "xxx", daily_paid: [12.3, 23.4, ...]},
      ...
    ]
  },
  zidahao: {
    summary: {                   // 自达号整体汇总
      直播GMV: 7330844.94,
      直播结算GMV: ...,
      消耗金额: ...,
      ROI: ...
    },
    sub_agencies: [],            // 5 个子机构汇总（花开/集米/太古/九三/直属自达号）
    anchors: [],                 // 138 个自达号达人
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
    daily_roi_by_sub: {          // 子机构每日加权 ROI
      "花开自达号": [10.5, 9.8, null, ...],
      ...
    },
    daily_roi_overall: [],       // 自达号整体每日 ROI
    anchor_detail: [             // 自达号下探：所有自达号达人分天支付明细（含 agency 字段）
      {name: "达人A", douyin_id: "xxx", agency: "花开自达号", daily_paid: [12.3, ...]},
      ...
    ]
  }
}
```

**关键**：`agency_top5_anchors` 是全部视图的字段名；`zidahao.top5_by_sub` 是自达号视图的字段名；`zidahao.anchor_detail` 包含 `agency` 字段用于子机构下探过滤。

## 前端功能

### 数据源切换

导航栏左侧新增「**全部** / **自达号**」切换按钮（橙色选中态），点击切换整个看板的数据源：

- **全部模式**（默认）：全局数据，3 张 KPI 卡片（GMV/结算/退款）、时间进度条、4 个图表、265 达人
- **自达号模式**：自达号专属数据，4 张 KPI 卡片（GMV/结算/消耗/ROI —— ROI 卡片可点击下钻）、**无时间进度**、3 个图表（总趋势/子机构趋势/子机构饼图）、138 自达号达人
- **自达号模式布局**：`agencyTrendCard` 通过 `grid-row: 1/3` 撑满右列，与左列两个卡片等高

切换逻辑由 `switchDataSource(source)` 统一管理，设置 `currentDataSource` 全局状态后销毁所有图表并重新渲染。

### 图表（全部模式）

- **达播组每日总业绩趋势**（折线图）: 总GMV、支付金额、退款金额，**近 30 天滚动窗口**，x 轴标签间隔显示（约 6 个）；**点击数据点可弹出全部达人的当日 vs 前日变化明细**（`showOverallDrill()`），不区分久酒/雅宁/星辞，跨机构列出所有达人，支持全部/仅下降/仅上涨筛选
- **每日机构业绩趋势**（折线图）: 7个主力机构，支持点击线/图例聚焦单个机构，再点恢复全部
- **久酒/雅宁/星辞分天业绩**（折线图）: 三人对比 + **「其他」**（灰色虚线 = 除三人外的全部达人）；标题右侧有 4 个按钮（全部/久酒/雅宁/星辞），点击聚焦单人，再点同一按钮恢复全部；**点击数据点可弹出该人员下所有达人的当日 vs 前日变化明细（人员下探）**（「其他」通过 `showOtherDrill()` 从 `all_anchor_daily` 中排除三人 douyin_id 后展示），支持全部/仅下降/仅上涨筛选
- **机构业绩占比**（饼图）: 7 个主力机构，仅统计当前月份（7月）业绩，从 `daily_by_agency` 中过滤 `dates` 以 `7/` 开头的日期索引求和，替代旧的全量 `pie_data`
- **人员业绩占比**（饼图）: 久酒/雅宁/星辞/其他 4 扇区，仅统计当前月份（7月）业绩，从 `person_daily` 中过滤 `7/` 开头的日期索引求和

### 图表（自达号模式）

- **自达号总业绩趋势**（折线图）: 自达号总 GMV/支付/退款；**点击数据点可弹出所有自达号达人的当日 vs 前日变化明细**（`showZidahaoDrill()`），支持全部/仅下降/仅上涨筛选
- **自达号子机构趋势**（折线图）: 5 条子机构线（花开/集米/太古/九三/直属自达号），支持点击聚焦；**点击数据点可弹出该子机构下达人的当日 vs 前日变化明细表**（`showZidahaoSubDrill()`），按变化量降序
- **子机构业绩占比**（饼图/环形图）: 5 个子机构扇区，仅统计当前月份（7月）业绩，从 `daily_by_sub` 过滤 `7/` 前缀日期求和（替代旧的全量 `pie_data`），点击弹出该子机构 Top5 自达号达人分天趋势弹窗
- **久酒/雅宁/星辞图表**：隐藏

子机构图表由 `renderZidahaoCharts()` 函数独立渲染，数据结构来自 `DATA.zidahao`。

**自达号 ROI KPI 下钻**：点击 ROI KPI 卡片（带「点击看趋势」提示），弹出各子机构每日 ROI 趋势折线图（5条子机构线 + 1条整体ROI紫色虚线，`showZidahaoRoiDrill()`）。数据来自 `zidahao.daily_roi_by_sub` 和 `zidahao.daily_roi_overall`，口径为 ΣGMV/Σ消耗 加权计算。<br>
**交互逻辑**：点击曲线 → 隐藏该机构（划掉），保留其他线；点击图例 → 切换对应的显示/隐藏；点击空白区域或「全部显示」按钮 → 恢复全部。

### 交互

- **数据源切换**: 点击「全部」/「自达号」按钮，销毁全部图表后重新渲染（`switchDataSource()`）
- **主题切换**: Header 右侧 🌙/☀️ 按钮，深色/浅色双主题，自动保存到 localStorage
- **机构趋势图**: 点击某条线只显示该机构，再点恢复全部（`window.__agencyFocusMode`）
- **三人业绩按钮**（`focusPerson()`）: 图表标题右侧 4 个 `.person-btn`，点击聚焦单人曲线，再点同一按钮恢复全部
- **人员下探**（`showPersonDrill(personName, dateIndex)`）: 点击三人图表任意数据点，弹出该人员下所有达人的当日 vs 前日变化明细表。按掉量降序（最大掉量排第一），红色 ↓ 标注下降、绿色 ↑ 标注上涨。顶部可切换「全部 / ↓ 仅下降 / ↑ 仅上涨」筛选。数据来自 `trend.person_anchor_detail`
- **「其他」下探**（`showOtherDrill(dateIndex)`）: 点击「其他」曲线任意数据点，从 `all_anchor_daily` 中排除久酒/雅宁/星辞三人下属达人的 douyin_id，得到「其他」达人集后展示当日 vs 前日变化明细表。复用人员下探弹窗和筛选按钮逻辑。纯前端过滤，无需后端改动
- **整体趋势下探**（`showOverallDrill(dateIndex)`）: 点击达播组每日总业绩趋势图任意数据点，弹出全部达人的当日 vs 前日变化明细表。复用人员下探弹窗和筛选按钮逻辑，按掉量降序排列。数据来自 `trend.all_anchor_daily`
- **自达号总业绩下探**（`showZidahaoDrill(dateIndex)`）: 点击自达号总业绩趋势图任意数据点，弹出所有自达号达人的当日 vs 前日变化明细表。数据来自 `zidahao.anchor_detail`
- **自达号子机构下探**（`showZidahaoSubDrill(subName, dateIndex)`）: 点击子机构趋势图数据点，弹出该子机构下达人的当日 vs 前日变化明细表。通过 `zidahao.anchor_detail` 的 `agency` 字段过滤
- **饼图下钻**: 点击机构扇区弹出 Top5 达人分天曲线（5 种颜色，支持弹窗内点击聚焦）。`showAgencyDrill()` 根据 `currentDataSource` 自动读取 `agency_top5_anchors`（全部）或 `top5_by_sub`（自达号）。饼图数据仅统计当前月份（7月），从 `daily_by_agency` / `person_daily` 按 `dates` 过滤 `7/` 前缀求和
- **达人下钻**: 点击达人列表中的「直播GMV」单元格弹出该达人每日支付曲线
- **ROI 下钻**（2 种入口）：
  - 达人表 ROI 列：点击弹出该达人每日 ROI 趋势（GMV/投放，琥珀色线条）—— `showRoiDrill()`
  - 自达号 ROI KPI 卡片：点击弹出各子机构每日 ROI 趋势（5 条线，加权口径）—— `showZidahaoRoiDrill()`
- **列拖拽排序**: 达人表头支持 HTML5 拖拽，拖动列头即可调整列顺序（`DISPLAY_COLS` / `ZDH_DISPLAY_COLS` 运行时可变）
- **列宽拖拽**: 鼠标移到表头右边缘，光标变 `col-resize`，拖拽调整列宽，宽度保存到 localStorage
- **搜索**: 输入框回车或点「查询」按钮，不实时过滤
- **GMV 目标**: 可自定义目标 GMV（万），进度条根据完成情况变色（绿/橙/红），保存在 `localStorage.targetGmvWan`（自达号模式下不显示）
- **应达 GMV 悬停**（`showExpectedGmv()` / `hideExpectedGmv()`）: 鼠标移到时间进度条上，下方弹出「当前时间进度下应达 GMV ≈ ¥XXX 万」

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
| `window.__agencyPieChart` / `window.__personPieChart` | 机构饼图 / 人员饼图实例 |
| `window.__drillChart` | 达人下钻弹窗图表实例（GMV / ROI 共用） |
| `window.__agencyDrillChart` / `window.__agencyDrillFocusMode` | 机构 Top5 达人趋势弹窗图表 / 聚焦状态 |
| `MAIN_AGENCIES` | 主力 7 机构：自达号、集米文化、紫语、花开满路、太古、亦初、直属 |
| `showPersonDrill(personName, dateIndex)` | 人员下探函数，弹出该人员所有达人的当日 vs 前日变化明细表 |
| `showOtherDrill(dateIndex)` | 「其他」下探函数，从 `all_anchor_daily` 排除三人后弹出，纯前端过滤 |
| `showOverallDrill(dateIndex)` | 整体趋势下探函数，弹出全部达人（跨机构）的当日 vs 前日变化明细表 |
| `showZidahaoDrill(dateIndex)` | 自达号总业绩下探函数 |
| `showZidahaoSubDrill(subName, dateIndex)` | 自达号子机构下探函数，按 agency 过滤 |
| `filterPersonDrill(filter)` | 人员下探筛选按钮（'all' / 'down' / 'up'） |
| `closePersonDrillModal()` | 关闭人员下探弹窗 |
| `window.__personDrillRows` / `__personDrillPrevDate` / `__personDrillClickedDate` | 下探弹窗缓存数据（共用） |
| `showExpectedGmv()` / `hideExpectedGmv()` | 时间进度条 hover：显示/隐藏应达 GMV 提示框 |

### 达人表显示规则

- `DISPLAY_COLS` / `ZDH_DISPLAY_COLS` 是运行时可变的数组（支持拖拽调整），通过 `getCols()` 根据 `currentDataSource` 切换
- **全部模式列**：`['主播昵称', '机构', '开播天数', '日均开播时长（小时）', '直播GMV', '直播退款GMV', '直播结算GMV', '结算率', 'ROI', '佣金支出', '投放消耗金额']`
- **自达号模式列**：将「佣金支出」替换为「时均 GMV」。时均 GMV = 直播GMV ÷ 开播天数 ÷ 日均开播时长（小时）
- **直播GMV 柱状条**（`.bar-cell` / `.bar-bg`）：绝对定位背景条，`opacity: 0.12`，宽度按当前列表最大 GMV 为 100% 比例计算，类似 Excel 数据条
- **ROI 列**：动态计算，无投放时显示 `-`
- **结算率**：保留颜色方块（绿 ≥40% / 橙 20-40% / 红 <20%）

## 部署上线

### 线上地址

| 平台 | 地址 | 说明 |
|------|------|------|
| GitHub Pages | `https://czcaizjy-lang.github.io/-/` | 主站，每次 git push 自动更新 |
| 本地预览 | `http://localhost:8976/dashboard.html` | 开发调试用，launchd 自动管理 |

### 自动部署（launchd WatchPaths）

Excel 保存 → macOS 内核检测 → `auto_sync.sh` 执行三步流水线 → 脚本退出，零残留。

```
Excel 保存
    │
    ▼  (macOS 内核 WatchPaths)
auto_sync.sh
    ├── ① sync_dashboard.py    → 生成 data/dashboard_data.json
    ├── ② build_standalone.py  → 生成 public/index.html
    └── ③ git push             → 推送到 GitHub Pages
```

### 手动部署

```bash
cd "/Users/xiaocao/CC/每日业绩自动统计"

# 第 1 步：提取数据
python3 scripts/sync_dashboard.py

# 第 2 步：构建独立页面（已由 sync_dashboard.py 自动调用，通常无需单独执行）
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

生成的页面无需本地服务器，可直接部署到任何静态托管平台。

### 部署架构

```
本地 Mac                              GitHub (czcaizjy-lang/-)
┌──────────────────┐                  ┌──────────────────────┐
│ Excel (.xlsx)    │                  │  main 分支            │
│      ↓           │                  │  ├── public/          │
│ sync_dashboard.py│    git push      │  │   └── index.html   │
│      ↓           │ ───────────────→ │  └── data/            │
│ build_standalone │                  │      └── dashboard_   │
│      ↓           │                  │         data.json     │
│ auto_sync.sh     │                  └──────────┬───────────┘
│ (launchd 触发)   │                             │
└──────────────────┘                   GitHub Pages (public/)
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

### 2. JSON 生成错误 / Sheet 找不到

脚本报 `KeyError: 'Worksheet xxx does not exist.'`

**解决**: 检查上游 Excel 文件的 Sheet 名称是否与脚本中硬编码的一致：
- `星辞业绩`
- `星辞自达号业绩`
- `X月直播数据`（自动识别含「直播数据」的 Sheet）
- `久酒业绩` / `雅宁业绩`

### 3. 日期类型错误

**原因**: Excel 中日期列可能是字符串而非 datetime 对象。

**解决**: 使用字符串解析：
```python
date_key = str(dt_val)[:10].replace('/', '-')  # 提取 "2026-06-14"
```

### 4. 抖音号类型

**原因**: 抖音号可能是字符串或数字。

**解决**: 始终用字符串比较：`douyin_id_str = str(douyin_id)`

### 5. 字段名错误

前端报 `Cannot read properties of null` 或 `undefined`。

**排查**: 用浏览器开发工具检查 JSON 字段名：
```javascript
console.log(DATA.trend.agency_top5_anchors);   // 机构 Top5
console.log(DATA.trend.person_daily);           // 三人分天支付
console.log(DATA.trend.person_anchor_detail);   // 人员下探达人明细
console.log(DATA.trend.all_anchor_daily);        // 全部达人下探明细
console.log(DATA.zidahao.anchor_detail);         // 自达号下探（含 agency 字段）
```

### 6. 自动同步不工作

**排查步骤**:

```bash
# 1. 检查 launchd 任务状态
launchctl list | grep dashboard
# 正常应显示：com.dashboard.sync（PID 为 - 表示休眠中）、com.dashboard.local-server（PID 有值）

# 2. 查看同步日志
tail -20 /Users/xiaocao/CC/每日业绩自动统计/data/sync.log

# 3. 检查 macOS 权限：系统设置 → 隐私与安全性 → 完全磁盘访问权限 → 确认 /bin/bash 已开启
#    如日志显示 "PermissionError: Operation not permitted"，说明权限未配置

# 4. 手动测试
bash /Users/xiaocao/CC/每日业绩自动统计/scripts/auto_sync.sh

# 5. 重载 launchd 任务
launchctl unload ~/Library/LaunchAgents/com.dashboard.sync.plist
launchctl load ~/Library/LaunchAgents/com.dashboard.sync.plist
```

### 7. 主题切换后图表颜色没变

主题切换时 `applyTheme()` 会调用 `renderTrendCharts()` 重绘所有图表。如果弹窗处于打开状态，关闭后重新打开即可获得新主题颜色。

### 8. 三个人分天业绩按钮不生效

- 确认 `focusPerson()` 函数存在（搜索 dashboard.html）
- 确认 `.person-btn` 按钮存在
- 打开浏览器控制台，检查 `window.__personChart` 是否已创建
- 强制刷新页面 (`Cmd+Shift+R`)，清除缓存

### 9. 浏览器缓存看到旧版本

- 按 `Cmd+Shift+R` 强制刷新
- 或访问 `http://localhost:8976/dashboard.html?v=2` 绕过缓存

### 10. 自达号模式切换后图表空白 / Canvas 报错

**原因**：Chart.js 不允许在同一 canvas 上创建第二个实例，必须先 `.destroy()`。

**解决**：`switchDataSource()` 已内置销毁逻辑。如出现，刷新页面即可。排查时检查 `currentDataSource` 变量值。

### 11. 自达号数据不显示或为 0

**可能原因**：
- `sync_dashboard.py` 未正确读取 `星辞自达号业绩` sheet
- JSON 中缺少 `zidahao` key

**排查**：
```javascript
console.log(DATA.zidahao);              // 检查 zidahao 数据是否存在
console.log(DATA.zidahao.summary);      // 汇总数据
console.log(DATA.zidahao.anchors.length);  // 应为 138
```

### 12. 自达号饼图点击没有弹出 Top5 达人

`showAgencyDrill()` 在自达号模式下从 `DATA.zidahao.top5_by_sub` 读取数据。如果某个子机构没有直播数据，会提示"暂无 Top 5 达人分天数据"。

### 13. 跨月数据更新

上游 Excel 路径中包含月份（`6月业绩`），跨月后需要更新：
- `sync_dashboard.py` 中的 `XLSX_PATH`
- `auto_sync.sh` 中的 `EXCEL_PATH`
- `~/Library/LaunchAgents/com.dashboard.sync.plist` 中的 `WatchPaths`
- 看板标题中的月份（`dashboard.html` h1）
- 建议每月初检查并更新路径

### 14. macOS 权限错误（PermissionError）

**现象**：`sync.log` 中显示 `PermissionError: [Errno 1] Operation not permitted`

**原因**：launchd 启动的进程无权访问桌面文件。WatchPaths 能检测到文件变化（内核级），但执行脚本时无法读取桌面路径。

**解决**：系统设置 → 隐私与安全性 → 完全磁盘访问权限 → 添加 `/bin/bash` 并开启。

## Git 托管

- **仓库**：`git@github.com:czcaizjy-lang/-.git`
- **线上看板**：`https://czcaizjy-lang.github.io/-/`（GitHub Pages）
- **分支**：main
- **SSH Key**：`~/.ssh/id_ed25519`（czcaizjy@gmail.com）
- **Git 配置**：user.name = xiaocao, user.email = czcaizjy@gmail.com
