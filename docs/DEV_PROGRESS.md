# 招标代理 SOP 工作台 — 开发进度与经验快照

> **最后更新**：2026-11-23（对应最新提交 `5b73491`）  
> **适用用途**：阶段性开发进度存档、业务规则备忘与跨会话上下文接续指南。

## 1. 项目概况与技术架构

- **项目目录**：`<workspace>/tender_workstation_gemini`
- **Git 仓库**：私密仓库 `silverashashash/tender-workstation-gemini`（`sample_docs/`、`.env`、SQLite 数据文件、`uploads/` 与 `*.tsbuildinfo` 均已加入 `.gitignore`，严禁将真实业务敏感文件提交至 Git）。
- **技术栈**：
  - **前端**：React 18 + TypeScript + Vite + Tailwind CSS v4 + Lucide 图标 + Dexie（IndexedDB 离线兜底）。
  - **后端**：Node.js（v24+ / v26+）+ Express + TypeScript + Node 原生 `node:sqlite`（`DatabaseSync`）+ Multer + SheetJS（`xlsx`）+ `exceljs` + `docx` + `pizzip`（原版 Word 模板 XML 保真填充）+ `archiver`。
  - **部署**：已配置 `Dockerfile` 与 `docker-compose.yml`，支持单用户密码锁（`AUTH_ENABLED=true` + `HttpOnly` 签名 Cookie）。
- **全局交互与视觉规范**：
  - 极简低饱和度办公风（`#F7F8FA` 页面底色、`#FFFFFF` 卡片、`#3B5B7E` 灰钢蓝主色；绿 `#ECFDF5/#2F6F5E`、琥珀 `#FEF9EC/#9A6B26`、红 `#FEF2F2/#A94442`）。
  - **左侧导航栏固定常显**（`sticky top-0 h-screen`）；**「2. 开评标模块」支持可折叠二级菜单**（6 个独立子板块，支持 `/evaluation/:section/:tool?` 深链），且核心 5 槽位上传矩阵在总览默认展开、其余子页默认折叠；工具型子板块采用**单工具聚焦视图**，便于后续持续扩展。
  - **全站统一分页**：所有数据明细表格统一采用 `20 项 / 50 项 / 全部` 分页组件（`PaginationBar.tsx`）。
  - **项目单位 / 业主单位**：全站统一为可搜索下拉选择框（`SearchableClientSelect.tsx`），对接 33 家客户协议库。
  - **全站日历与 24 小时制时间选择**（`DateTimePickerField.tsx`）：所有日期均通过日历选择；具体时间设为可选，且采用 **24 小时制（00–23 时 / 00–59 分）** 下拉选择，不受系统语言 AM/PM 影响。

## 2. 模块开发进度总览

| 模块 | 路由 | 状态 | 核心实现内容 | 关键代码文件 |
| :--- | :--- | :--- | :--- | :--- |
| **工作台总览** | `/` | 已完成 v1.0 | 多批次切换、全流程进度卡片、SOP/归档完成度统计 | `client/src/pages/DashboardPage.tsx`、`client/src/components/Layout.tsx` |
| **登录锁** | `/login` | 已完成 v1.0 | 公网部署单用户访问密码保护 | `client/src/pages/LoginPage.tsx`、`server/src/auth.ts` |
| **1. 招标模块** | `/bidding` | **已完成 v2.2** | 5 个 SOP 槽位（均支持一键「清除」与重建）；双 ERP（主业 SAP/ECP + 产业 ETP/ECP）自适应解析与联合采购追加上传；主键互通（`采购申请唯一标识`/`采购申请标识`/`采购申请号`/`采购申请` 10 位↔15 位）；以《需求一览表》为权威基准过滤表外项并持久化冲突预警；全量差异三表分列 Excel 导出（数量差异置顶）；精简/全量视图、折扣率、权重纠偏、围串标初筛、拼音+包号排序、批量删除 | `server/src/biddingParser.ts`、`server/src/routes/bidding.ts`、`client/src/pages/BiddingPage.tsx`、`client/src/utils/projectSort.ts` |
| **2. 开评标模块** | `/evaluation/:section?` | **已完成 v3.1** | **6 大二级子板块**：① **评审总览与数据**（5 槽位上传矩阵、跨批次防错传拦截、同标包聚合、流标预警与彩色风险标签）；② **报价与税率核验**（超限价、报价过低、二轮超一轮、相同报价、税率异常/缺失/证明待核实，分表导出报价/税率两份 Excel）；③ **围串标风控核验**（同标包联系信息重合 + 投标文件硬件特征 CPU/硬盘/MAC 全一致预警及导出）；④ **现场核验与算分**（全量迁移 6 大工具：专家打分结果校验、价格分计算、履约分生成、经济排序表核验、授标分配、智能分包）；⑤ **现场汇报与报告**（已接入预警事项确认单、初评汇报生成、初评会议纪要生成）；⑥ **SOP 与阶段附件** | `server/src/evaluationParser.ts`、`server/src/routes/evaluation.ts`、`server/src/routes/toolboxTools.ts`、`server/src/toolboxWorkbook.ts`、`server/src/expertScoreAudit.ts`、`server/src/priceScoreTool.ts`、`server/src/performanceTool.ts`、`server/src/economicRankTool.ts`、`server/src/awardTool.ts`、`server/src/splitPackTool.ts`、`client/src/pages/EvaluationPage.tsx` |
| **3. 定标模块** | `/award` | **已完成 v2.2** | 6 大定标交付物一键生成（预成交/中标候选人公示 Word、成交/中标结果公告 Word、成交通知书交接单 Excel、项目交接信息表 Excel、成交结果表 Excel、服务费计算表 Excel）；精确分段累进服务费与 33 家客户协议折扣引擎；ECP 供应商联系人按「供应商+分标名称+包号」三级精确匹配；多原因否决规范拼接（`商务否决：xxx；技术否决：yyy`）；13,767 条物料编码库自动匹配物料描述与 `.xlsx` 在线更新；公示/公告日期常量联动回填 | `server/src/routes/award.ts`、`server/src/awardGenerator.ts`、`server/src/serviceFee.ts`、`server/src/materialCatalog.ts`、`client/src/pages/AwardPage.tsx` |
| **4. 资料整理模块** | `/archive` | 待阶段 4 实表细化 | 双 Sheet 归档目录导入（B 列底色自动标必交）、`已有/缺失/不涉及(无)` 三态跟踪、跨模块文件关联、规范格式台账 Excel 与 ZIP 打包导出 | `server/src/routes/archive.ts`、`client/src/pages/ArchivePage.tsx` |

## 3. 开评标模块 v3.0 核验与工具清单

### 3.1 报价与税率核验、围串标风控规则
1. **报价核验统一口径**：
   - **整包各轮报价**：以《开标一览表》为主（未上传时用《行信息报价基础数据表》兜底）。
   - **合并子包各轮报价**：以《行信息报价基础数据表》按子包逐轮独立核验。
   - **4 项报价检查**：`超限价`、`报价过低`（低于成本 3 项规则：低于初评均值 50%、低于次低价 50%、低于最高限价 45%）、`二轮超一轮`（否决项）、`不同供应商相同报价`。
   - **2 份独立导出**：`报价核验 (.xlsx)`（汇总超限价、报价过低、二轮超一轮、相同报价）与 `税率核验 (.xlsx)`（汇总税率异常、税率缺失、税率证明待核实）。
   - **行内风险标签位置**：报价类彩色标签紧贴最终报价列，税率异常/缺失标签紧贴增值税率列，初评状态列仅保留「正常 / 否决」。
2. **围串标风控核验**：
   - **联系信息重合**：严格限定同分标同包不同供应商，比对联系人、手机、固话、邮箱、传真、邮政地址（邮政地址须含数字或“号”）。
   - **硬件特征核验**（第 3 槽位《投标文件IP地址一览表》）：同标包跨文件类别、跨轮次比对，仅 `CPU + 硬盘 + MAC` 全部一致触发围串标预警；缺 CPU 或硬盘但含其他字段列为「资料不足，未完成核验」。

### 3.2 已迁移 / 新开发的现场工具
| 所属子板块 | 工具名称 | 来源 | 核心特性与输出 |
| :--- | :--- | :--- | :--- |
| 现场核验与算分 | **专家打分结果校验**（总分 / 客观分） | 迁移自工具箱 工具 8 & 9 | 校验专家评分汇总一致性与同包客观分偏差，支持结果持久化与导出 (`ExpertScoreCheckPanel.tsx`) |
| 现场核验与算分 | **价格分计算** | 迁移自工具箱 `PriceScoreTool` | 完整实现国网 6 种基准价与价格分公式（基于 `BigInt` 高精度十进制 `Dec` 类），支持总价/折扣率、轮次筛选、6 种公式说明与标准模板下载 (`PriceScorePanel.tsx`) |
| 现场核验与算分 | **履约分生成** | 迁移自工具箱 `PerformanceTool` | 支持「供应商」「供应商+标段」「供应商+所属部门」三类关联，严格保留 1–5 分映射、边界 30 及多条匹配末条有效/默认 4 分复核提示 (`PerformanceScorePanel.tsx`) |
| 现场核验与算分 | **经济排序表核验** | 迁移自工具箱 `EconomicRankTool` | 先排查技术/商务/价格分为 0 异常，再按「总分降序 → 技术分降序 → 投标总价升序 → 商务分降序」复核排序一致性 (`EconomicRankPanel.tsx`) |
| 现场核验与算分 | **授标分配** | 迁移自工具箱 `AwardTool` | 支持无限授、一轮一包、一轮多包三种模式，兼容原版经济排序表（完整/精简）与标准 6 列排序表，支持限授组配置与并列复核提示 (`AwardAllocPanel.tsx`) |
| 现场核验与算分 | **智能分包** | 迁移自工具箱 `SplitPackTool` | 严格复刻三阶段数量拆分守恒算法（保底单位 → 冻结比例切分 → 高单价轮询兜底），支持按比例/按金额及重量/长度粒度控制，导出汇总+明细双表 (`SplitPackPanel.tsx`) |
| 现场汇报与报告 | **预警事项确认单** | 迁移自工具箱 工具 1 | 严格沿用原 Java 逻辑与原版 Word 模板（`server/templates/预警事项确认单.docx`），保真生成确认单 (`WarningFormPanel.tsx`) |
| 现场汇报与报告 | **初评汇报生成** | 迁移自工具箱 工具 6 | 读取《初评情况汇报模板》附件 Sheet（跳过封面 1/2），生成 Word 初评汇报单，正文末尾固定追加段落「汇报完毕！」(`InitialReportPanel.tsx`) |
| 现场汇报与报告 | **初评会议纪要生成** | 全新开发 | 内置「公开」「非公开」双原版 Word 模板（自动切换评标/评审委员会、引用章节及表格列名）；动态替换批次名称、开会时间、评审组别与澄清说明并自动清除黄色高亮；支持智能表头行扫描（按别名优先级匹配，兼容《初评情况汇报模板》第 2 行表头、《未进入详评情况统计表》第 4 行表头及自定义表）与商务/技术分组筛选 (`InitialMinutesPanel.tsx`) |

## 4. 核心业务规则与导出格式规范

### 4.1 术语、平台与联合采购规则
1. **公开招标 vs 非公开招标**：
   - 公开招标：`招标人 / 投标人 / 招标文件 / 投标文件`，单轮报价（开标须做价格公示），进入详评需 **≥ 3 家**。
   - 非公开招标：`采购人 / 响应供应商 / 采购文件 / 响应文件`，支持多轮报价，进入详评需 **≥ 2 家**。
2. **主业 ERP vs 产业 ERP 与联合采购**：
   - 联合采购批次可同时包含主业与产业项目（统一在 ECP 执行），`erp_plan` 槽位支持分次追加上传主业与产业两份 ERP 采购计划并自动合并；若批次已在 ECP 执行，后传产业表不降级平台。
   - 主键互通：`采购申请唯一标识`（主业 15 位 = 10 位采购申请 + 5 位行项目号如 `00010`）、`采购申请`（主业 10 位）、`采购申请标识`（产业 19 位）、`采购申请号`（需求一览表）统一互通匹配。
3. **《需求一览表》权威基准与 ERP 差异校验**：
   - 《需求一览表》是最终挂网方案；一旦上传，项目数量与核心字段严格以《需求一览表》为准，自动过滤前期申报表或 ERP 中多出的表外项目，后上传的 ERP/申报表仅做补充字段丰富，不覆盖核心字段。
   - 自动比对数量差异（`ERP缺失`、`ERP多出`）与 5 项关键字段冲突（`项目名称`、`项目概况`、`预算金额`、`最高限价`、`技术规范书编码`，文本忽略中英文标点与空格），在页面置顶展示可筛选预警面板，并支持一键导出三工作表全量差异 Excel。

### 4.2 定标匹配、否决拼接与物料描述规则
1. **ECP 供应商联系人精确匹配**：按 `供应商名称 + 分标名称 + 包号` → `供应商名称 + 包号` → `供应商名称`（兼容 ETP）三级匹配，避免同一供应商投多包填不同联系人时串人。
2. **多原因否决规范拼接**：《成交结果表》备注栏将同一供应商在同一包内的多条否决记录去重并按固定顺序拼接为 `商务否决：xxx；技术否决：yyy`（流标包记为 `流标（商务否决：xxx；技术否决：yyy）`）。
3. **物料编码 → 物料描述自动匹配**：启动时从 `sample_docs/3.award/其他文件/服务物料编码--分标名称（授权采购）.xlsx` 自动导入 13,767 条对照库至 `material_catalog` 表（支持页面上传 `.xlsx` 在线更新）。
4. **日期常量联动**：公示与公告卡片提供持久化的 `成交公示日期` 与 `成交公告日期` 日历选择框，自动联动回填《项目交接信息表》对应日期列。

### 4.3 Word & Excel 统一导出规范
- **Word（公示 / 公告）**：表格单元格水平+垂直居中、表头无底纹、紧凑行高（`25 dxa` 上下边距、`0` 段距、9pt 宋体）、跨页自动重复表头（`<w:tblHeader/>`）、禁止单行跨页拆断（`<w:cantSplit/>`）、价格去除末尾多余 `0`；支持按需勾选 4 个可选列。
- **Word（现场报告原版模板）**：使用 `pizzip` 直接操作原版 `.docx` 的 `word/document.xml`，完整保留原始页眉页脚、红头/字号与段落样式，同时自动清除 `<w:highlight>` 与黄色底纹。
- **Excel（全工作台所有导出）**：统一纯黑（`#000000`）**Calibri 11 号字体**（表头加粗无底色）、全部单元格（含空单元格）细黑全边框（All Borders）、表头开启自动筛选（AutoFilter）、打印重复表头（`printTitlesRow`）、紧凑行高；《成交通知书交接单》包含流标项目。

## 5. 数据库表结构速查（`server/src/db.ts`）

- **基础与招标**：`batches`（含 `award_*` 定标配置、`erp_conflicts_json` 差异缓存、`eval_audit_config_json` 税率口径配置）、`batch_projects`（含 `material_code` / `material_desc`）、`batch_suppliers`（含 `package_no`）、`uploaded_files`
- **开评标与现场工具**：`eval_bids`、`eval_line_quotes`、`eval_ip_hardware`、`eval_rejections`、`eval_rankings`、`eval_score_checks`（通过 `check_type` 统一持久化 `total_score` / `objective_score` / `initial_report` / `warning_form` / `initial_minutes`）、`sop_items`
- **定标与基础库**：`award_outputs`、`client_agreements`（33 家协议库）、`material_catalog`（13,767 条物料编码库）
- **归档**：`archive_items`

## 6. 关键踩坑与工程经验总结

1. **ESM 下 SheetJS 读取**：`xlsx` v0.18.5 在 ESM 中必须使用 `(XLSXModule as any).default ?? XLSXModule` 并通过 `XLSX.read(fs.readFileSync(filePath), { type: 'buffer' })` 读取。
2. **动态表头匹配必须按别名优先级遍历**：例如《初评情况汇报模板》同时包含「问题描述」（C 列）和「未进入详评原因」（F 列），若按列从左到右先到先得会误匹配「问题描述」；改为先遍历别名优先级数组再查找对应列即可精准锁定主列。
3. **原版 Word 模板跨 Run 拆分问题**：Word 文档中单个中文词组常被拆分为多个 `<w:r><w:t>` 节点，定位或校验段落时应先拼接段落纯文本，替换时复用首个 `<w:rPr>` 并移除黄色高亮标签，既保真版式又避免残留标黄。
4. **主业 SAP 15 位主键与需求表 10 位主键互通**：通过 `isSameReqAppNo` 兼容 10 位↔15 位前缀互通，在 `12EJ02` 联合采购实测实现 **951/951（100%）** 命中。
5. **ECP 多联系人按标包分流**：`12EJ02` 实测有 51 家供应商在不同标包填写了不同联系人，必须按 `分标包号` 展开三级索引。
6. **Git 推送与构建**：
   - 本地构建：`export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH" && npm run build --prefix server && npm run build --prefix client`
   - 推送 GitHub：因沙盒环境无法读取 macOS Keychain 凭据，提交后在本机终端执行 `git push origin main`。

## 7. 下一步计划

1. **继续扩展开评标现场报告模板**：
   - 「现场汇报与报告」待接入：法律报告生成、现场监督报告生成、评审报告生成。
2. **阶段 4（资料整理归档模块 `/archive`）**：结合真实归档清单模板与实操场景细化自动匹配与打包归档流程。
