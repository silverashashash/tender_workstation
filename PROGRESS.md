# 招标代理 SOP 工作台 — 项目接力开发进度（PROGRESS）

> **最后更新**：2026-11-24（对应最新代码提交 `4d74b01`）  
> **技术栈**：React 18 + Vite + Tailwind CSS v4 + TypeScript（前端）｜Node.js (v22+/v24+) + Express + 原生 `node:sqlite` + ExcelJS / SheetJS / PizZip（后端）  
> **详细业务规则与历史踩坑记录**：另见 [docs/DEV_PROGRESS.md](docs/DEV_PROGRESS.md)

## 1. 当前目录结构

```text
tender_workstation_gemini/
├── PROGRESS.md                     # 本文件：跨设备接力开发总览
├── PRD.md                          # 产品需求文档
├── package.json                    # 根目录统一脚本（install:all / dev / build / start）
├── Dockerfile                      # 生产镜像构建配置
├── docker-compose.yml              # Docker 一键部署配置
├── docs/
│   └── DEV_PROGRESS.md             # 详细业务规则、数据库说明与历史踩坑文档
├── sample_docs/                    # [已 gitignore] 真实业务样本文件，不入库
├── server/
│   ├── data/workstation.db         # [已 gitignore] SQLite 业务数据库（可通过 DATA_DIR 指定同步路径）
│   ├── uploads/                    # [已 gitignore] 上传原表与生成交付物文件
│   ├── templates/                  # 内置原版 Word 模板（随代码入库）
│   │   ├── 预警事项确认单.docx
│   │   ├── 初评会议纪要_公开.docx
│   │   └── 初评会议纪要_非公开.docx
│   └── src/
│       ├── index.ts                # Express 入口（默认端口 3001）
│       ├── auth.ts                 # 单用户密码锁与 Cookie 鉴权
│       ├── db.ts                   # node:sqlite 建表与迁移
│       ├── seed.ts                 # 初始演示数据与 33 家客户协议库初始化
│       ├── materialCatalog.ts      # 13,767 条服务物料编码对照库导入与查询
│       ├── serviceFee.ts           # 分段累进代理服务费与客户协议折扣引擎
│       ├── excelParser.ts          # 通用表格预览解析
│       ├── biddingParser.ts        # 招标 5 类表解析（双 ERP、需求一览表、权重纠偏、概算金额/技术规范书ID 别名）
│       ├── evaluationParser.ts     # 开评标 5 槽位解析（开标一览表、行报价、IP硬件、否决、排序）
│       ├── evaluationAudit.ts      # 报价核验、税率核验、围串标风控
│       ├── awardGenerator.ts       # 定标 6 大交付物生成（Word 公示/公告 + 4 份规范 Excel）
│       ├── docGenerator.ts         # 通用文档生成工具
│       ├── expertScoreAudit.ts     # 专家打分校验（总分 + 客观分）
│       ├── toolboxWorkbook.ts      # 工具箱高精度十进制 Dec(BigInt) 与通用 Excel 读写底层
│       ├── priceScoreTool.ts       # 价格分计算（国网 6 种公式）
│       ├── performanceTool.ts      # 履约分生成（3 类关联 + 边界 30 复核）
│       ├── economicRankTool.ts     # 经济排序表核验
│       ├── awardTool.ts            # 授标分配（无限授 / 一轮一包 / 一轮多包）
│       ├── splitPackTool.ts        # 智能分包（三阶段数量守恒算法）
│       ├── warningFormGenerator.ts # 预警事项确认单 Word 生成
│       ├── initialReportGenerator.ts  # 初评汇报单 Word 生成（结尾含「汇报完毕！」）
│       ├── initialMinutesGenerator.ts # 初评会议纪要 Word 生成（公开/非公开双模板 + 动态表头扫描）
│       └── routes/
│           ├── batches.ts          # 批次 CRUD 与项目增删改
│           ├── bidding.ts          # 招标 5 槽位上传/清除、ERP 差异校验与导出、增量导入
│           ├── evaluation.ts       # 开评标上传、核验、导出与现场报告接口
│           ├── toolboxTools.ts     # 价格分/履约分/经济排序/授标分配/智能分包 API
│           ├── award.ts            # 定标数据聚合、6 大交付物生成、分页/批量删除
│           ├── archive.ts          # 资料整理归档目录与打包导出
│           ├── sop.ts              # 各阶段 SOP 清单与附件
│           ├── files.ts            # 文件下载与预览
│           └── backup.ts           # 数据备份与恢复
└── client/
    └── src/
        ├── App.tsx                 # 路由配置（含 /evaluation/:section/:tool? 深链）
        ├── api/client.ts           # 全站后端 API 封装
        ├── context/BatchContext.tsx # 全局当前批次切换与刷新上下文
        ├── types/index.ts          # 前后端共享 TypeScript 类型
        ├── utils/
        │   ├── terminology.ts      # 公开/非公开 & 主业(ECP)/产业(ETP) 术语与金额格式化
        │   └── projectSort.ts      # 标段拼音 + 自然包号排序
        ├── pages/
        │   ├── DashboardPage.tsx   # 工作台总览（批次切换与进度卡片）
        │   ├── LoginPage.tsx       # 访问密码锁登录页
        │   ├── BiddingPage.tsx     # 1. 招标模块
        │   ├── EvaluationPage.tsx  # 2. 开评标模块（6 大子板块 + 单工具聚焦视图）
        │   ├── AwardPage.tsx       # 3. 定标模块
        │   ├── ArchivePage.tsx     # 4. 资料整理模块
        │   └── ToolboxPage.tsx     # 实用工具箱（跨阶段独立工具卡片墙）
        └── components/
            ├── Layout.tsx                  # 固定左侧导航栏 + 顶部可搜索批次切换栏
            ├── SearchableBatchSelect.tsx   # 顶部可搜索批次选择器（编号前置 + 编号/名称多词搜索）
            ├── DateTimePickerField.tsx     # 统一日历 + 24 小时制时间选择
            ├── SearchableClientSelect.tsx  # 33 家协议单位可搜索下拉框
            ├── PaginationBar.tsx           # 全站统一分页组件
            ├── ExcelPreviewModal.tsx       # 在线 Excel 原表预览弹窗
            ├── ExpertScoreCheckPanel.tsx   # 专家打分校验面板
            ├── PriceScorePanel.tsx         # 价格分计算面板
            ├── PerformanceScorePanel.tsx   # 履约分生成面板
            ├── EconomicRankPanel.tsx       # 经济排序表核验面板
            ├── AwardAllocPanel.tsx         # 授标分配面板
            ├── SplitPackPanel.tsx          # 智能分包面板
            ├── WarningFormPanel.tsx        # 预警事项确认单面板
            ├── InitialReportPanel.tsx      # 初评汇报生成面板
            └── InitialMinutesPanel.tsx     # 初评会议纪要生成面板
```

## 2. 已完成模块

| 模块 | 路由 | 当前版本 | 核心已完成功能 |
| :--- | :--- | :--- | :--- |
| **工作台总览 & 登录锁** | `/`、`/login` | v1.0 | 多批次切换、全流程进度卡片、SOP/归档完成度统计；支持单用户密码锁 |
| **1. 招标模块** | `/bidding` | **v2.3** | 5 个 SOP 文件槽位（支持一键清除与重建）；主业 SAP/ECP + 产业 ETP/ECP 双 ERP 联合采购追加合并；10 位↔15 位主键互通；以《需求一览表》为权威基准过滤表外项并导出三表差异 Excel；折扣率/合并成包识别、权重纠偏、围串标初筛；已兼容 `技术规范书编码`↔`技术规范书ID` 别名，及 `概算金额（含税,万元）`、`资质要求`、`工程进度安排`、`合同名称`、`物料编码/描述` 等新表头 |
| **2. 开评标模块** | `/evaluation/:section/:tool?` | **v3.1** | 左侧栏可折叠二级菜单 + 6 大子板块：评审总览与数据、报价与税率核验（分表导出报价/税率两份 Excel）、围串标风控核验（联系信息 + IP 硬件特征）、现场核验与算分（6 工具：专家打分校验、价格分计算、履约分生成、经济排序表核验、授标分配、智能分包）、现场汇报与报告（3 工具：预警事项确认单、初评汇报生成、初评会议纪要生成）、SOP 与阶段附件 |
| **3. 定标模块** | `/award` | **v2.3** | 6 大定标交付物一键生成（公示 Word、结果公告 Word、成交通知书交接单、项目交接信息表、成交结果表、服务费计算表）；33 家客户协议折扣与分段累进服务费引擎；ECP 三级联系人精确匹配；多原因否决规范拼接；13,767 条物料编码库在线更新；公示/公告日期联动回填；已生成交付文件记录支持每页 5 份分页展示与勾选批量删除 |
| **4. 资料整理模块** | `/archive` | v1.0（基础版） | 双 Sheet 归档目录导入、`已有/缺失/不涉及(无)` 三态跟踪、跨模块文件关联、台账 Excel 与 ZIP 打包导出 |

## 3. 待改项与后续计划

1. **开评标 →「现场汇报与报告」待接入的报告工具**：
   - [ ] 法律报告生成
   - [ ] 现场监督报告生成
   - [ ] 评审报告生成
   - [ ] 其余待迁移的现场表格/报告工具
2. **阶段 4：资料整理模块（`/archive`）实表细化**：
   - [ ] 结合真实归档清单模板完善自动匹配规则、缺失项提醒与按目录结构打包 ZIP。
3. **跨设备数据同步说明**：
   - `server/data/workstation.db`、`server/uploads/` 与 `sample_docs/` 已加入忽略列表，不随 Git 同步。
   - 如需在另一台电脑复用同一份批次数据，可在本地环境配置中将 `DATA_DIR` 指向同一个云盘目录（如 iCloud Drive），或通过工作台备份/恢复功能迁移。
4. **遗留轻量数据说明**：
   - 早期内置演示批次 `GW-AH-2026-0412` / `0418` / `0425` 存在未填 `budget_yuan` 的种子记录，不影响真实上传批次；`CY1226SDD006` 已在本地库回填完成（含税预算合计 `544.1334 万元`）。

## 4. 运行、构建与同步命令

> **环境要求**：Node.js v22.10+（推荐 v24+ / v26+，后端依赖原生 `node:sqlite`）。

### 4.1 首次在新电脑安装依赖
```bash
npm run install:all           # 同时安装根目录、server、client 依赖
```

### 4.2 本地开发启动（前后端热更新）
```bash
npm run dev
```
- 前端 Vite 开发服务：`http://localhost:5173`
- 后端 Express API 服务：`http://localhost:3001`

### 4.3 生产构建与启动
```bash
# 全量构建（前端 dist + 后端 dist）
npm run build

# 单独构建后端或前端
npm run build --prefix server
npm run build --prefix client

# 启动生产服务（后端自动托管前端静态产物，默认 3001 端口）
npm run start

# Mac mini 方案一：一键编译并注册/重启 macOS 后台常驻服务（开机自启 + Tailscale 直连）
npm run mac:deploy
# 拉取 GitHub 最新代码并自动重新编译重启：
npm run mac:deploy -- --pull

# 一键生成 Windows 内网绿色便携包（输出至 release/tender-workstation-win-portable/）
npm run build:win-portable
# 如需连同当前本机数据库与附件一起打包带进现场：
npm run build:win-portable -- --include-data
# Windows 源码目录直接双击傻瓜启动脚本：scripts/win-setup-and-start.bat
```

### 4.4 Docker Compose 部署
```bash
docker compose up -d --build
```

### 4.5 跨设备 Git 同步
```bash
# 切换设备前推送代码
git status
git push origin main

# 在另一台电脑拉取最新代码
git pull --rebase origin main
npm run install:all   # 仅当 package.json 有新增依赖时需要执行
```
