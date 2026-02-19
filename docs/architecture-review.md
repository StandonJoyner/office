# Office Suite 架构设计评审报告

> 评审视角：资深架构师  
> 对照文档：`requirements.md`（需求） vs `architecture.md`（架构）  
> 结论：**架构整体合理且与需求基本匹配**，存在若干需补齐或统一之处。

---

## 1. 总体结论

| 维度         | 结论 |
|--------------|------|
| 需求覆盖度   | 高（P0/P1 核心需求均有对应设计） |
| 分层与职责   | 清晰（Presentation → Editor Core → Business → Data） |
| 同源与扩展   | 适配器 + 事件总线 + 插件，可支撑 Web/桌面与扩展 |
| 待改进点     | 术语统一、Office 格式、协作与性能可测性、冲突策略 |

---

## 2. 需求与架构对照

### 2.1 核心价值主张（需求 1.2）✅

| 需求           | 架构对应 |
|----------------|----------|
| 数据单一事实来源 | Excel Data Manager + DataLinkManager 的 source/target 模型 |
| 动态更新       | SyncEngine（手动/自动刷新、冲突检测） |
| 完整数据溯源   | Data Traceability Service（版本、历史、影响分析） |
| Web/桌面同源   | 适配器模式（Storage/File/Notification）+ 4.3 构建策略 |

**结论**：四大价值主张在架构中均有明确落点。

---

### 2.2 引用类型与操作（需求 3.1.2）✅ 含一处不一致

| 需求项       | 架构对应 | 说明 |
|--------------|----------|------|
| 单元格引用   | `ReferenceType`: cell | 一致 |
| 范围引用     | range | 一致 |
| 公式引用     | formula | 3.2 节有，2.2 节 DataReference 仅写 `'cell' \| 'range' \| 'named'`，**需统一为含 `formula`** |
| 命名区域     | named | 一致 |
| 插入/显示/跳转/断链 | DataLinkManager CRUD + ReferenceDisplay + TraceabilityService | 跳转、断链修复流程建议在架构中单列一小节 |

**建议**：在 2.2 节将 `DataReference.type` 与 3.2 的 `ReferenceType` 统一为 `'cell' | 'range' | 'named' | 'formula'`，并在 2.2 或 3.2 中简要说明「断链检测与修复」的职责归属（如 DataLinkManager + UI）。

---

### 2.3 数据更新机制（需求 3.1.3）✅ 小缺口

| 需求项       | 架构对应 |
|--------------|----------|
| 手动刷新     | `refreshReference` / `refreshDocument` / `refreshAll` ✅ |
| 自动刷新     | `enableAutoSync` / `setSyncInterval` ✅ |
| 选择性刷新   | 仅有按引用 id、按文档 id、全部；**缺少「按数据源/按区域」** |
| 更新冲突解决 | `detectConflicts` / `resolveConflict(strategy)` ✅，策略未展开 |

**建议**：
- 在 SyncEngine 接口中增加「按源刷新」能力，例如：`refreshBySource(source: DataSource): Promise<void>`，与需求「按引用源或区域更新」对齐。
- 在 2.3 或单独小节简述 `ResolveStrategy` 的几种策略（如：保留源、保留文档、手工选择、最后写入胜出），便于实现与验收。

---

### 2.4 Excel 数据源与变更追踪（需求 3.2）✅

| 需求项           | 架构对应 |
|------------------|----------|
| 文件/工作表管理   | ExcelDocument + Univer Workbook/Worksheets |
| 数据变更追踪     | Custom API: CellChangeTracker、ValueSnapshot |
| 影响分析（受影响 Word 引用） | TraceabilityService.getDependents(cellId): ReferenceId[] |

**结论**：需求 3.2 在编辑器核心与溯源服务中均有体现。

---

### 2.5 数据溯源（需求 3.3）✅

| 需求项         | 架构对应 |
|----------------|----------|
| 引用链追踪     | TraceabilityService + DocumentSnapshot.references |
| 版本管理       | createSnapshot / listSnapshots / compareSnapshots |
| 影响分析（正/反向） | getDependents / getDependencies |
| 审计日志       | logAction / getAuditLog |

**结论**：溯源需求覆盖完整。

---

### 2.6 技术栈与性能（需求 4）⚠️

| 需求项           | 架构对应 | 说明 |
|------------------|----------|------|
| Next.js / Tiptap / Univer / Zustand·React Query / Tauri | 5.1–5.4 技术栈表 | 一致 |
| 文档打开 <2s     | 6.1 编辑器性能（虚拟滚动、延迟加载、增量更新） | 有手段，**缺可验证的指标或测试策略** |
| 刷新 <500ms      | 6.2 缓存、批量刷新、异步解析 | 同上 |
| 10,000+ 行 Excel | 6.2–6.3 | 未显式写「大表」策略（如按需加载、视口内计算） |
| 50+ 并发用户     | 未在架构中体现 | 涉及实时协作与后端，当前架构偏单机/多实例；若 P2 才做协作，可接受，**建议在路线图中标明** |

**建议**：
- 在 6.1–6.3 中补充一句「性能验收与负载策略」（如：基准测试、目标 2s/500ms、10k 行场景）。
- 在 10 开发路线图中注明「50+ 用户」依赖 P2 协作/后端，避免与当前单机架构混淆。

---

### 2.7 安全、兼容性、可扩展（需求 5）✅ 含一处缺口

| 需求项           | 架构对应 |
|------------------|----------|
| 数据加密/传输/审计 | 7.1 数据安全 + 7.2 访问控制 + TraceabilityService 审计 | ✅ |
| Office 格式兼容   | **未在架构中单独列出** | 需求 5.2：.docx/.xlsx；当前为 Tiptap/Univer 原生 JSON + document.office 存储 |
| 插件与扩展       | 8.1 插件系统 + 8.2 事件总线 | ✅ |

**建议**：在「数据层」或「存储格式」旁增加一小节「Office 格式兼容策略」：导入/导出 .docx/.xlsx 的职责（转换层/服务）、与现有 Tiptap/Univer JSON 的关系、是否纳入 P1（文件系统支持）或 P2。

---

## 3. 架构图与模块命名统一

### 3.1 图示与正文不一致

- 架构图 1.1 中为：**Document Store Manager**、**Reference Service**、Excel Data Manager、Sync Engine。
- 正文 2.2 核心描述为 **Data Link Manager**，且与代码/CLAUDE 一致。
- 「Document Store Manager」「Reference Service」在正文中无接口定义与职责划分。

**建议**：
- 在 1.1 图旁增加图例或一句话：  
  「*Reference Service 即 Data Link Manager（数据链接管理器）；Document Store Manager 负责文档/工作簿的持久化与元数据。*」
- 或在 2.2 前增加「Document Store Manager」的职责说明（如：文档 CRUD、版本号、与 Data Layer 的交互），避免读者与 DataLinkManager 混淆。

### 3.2 DataReference 两处定义

- **2.2**：`source: { fileId, sheetId, range }`，结构简略。
- **3.2**：`source: DataSource`（含 fileName, sheetName, range, isFormula 等），与 `core/types.ts` 一致。

**建议**：2.2 改为「source: DataSource（见 3.2）」或直接引用 3.2，只保留一份权威定义，避免实现时两套模型。

---

## 4. 未在架构中展开但需求存在的点

| 需求/开放问题           | 建议 |
|-------------------------|------|
| 协作编辑（多人、评论、修订） | 需求 3.1.1 / P2。在 8 扩展架构或 10 路线图中增加「协作层」扩展点（如 CRDT/OT、Presence、Comment 模型），标明 P2。 |
| 离线模式                 | 需求 8 开放问题。可在 4.2 适配器或 Data Layer 中增加「离线/同步」扩展说明（如 Service Worker + 队列）。 |
| 大文件分块加载           | 需求 8。在 6.3 或 3.3 存储格式旁补充「大文件分块与按需加载」策略。 |
| 协作冲突解决详细策略     | 需求 8。与 2.3 的 ResolveStrategy 合并，在架构中列出一小节「协作与更新冲突策略」。 |

---

## 5. 架构优点（保持）

1. **分层清晰**：Presentation → Editor Core → Business → Data，边界明确，利于 Web/桌面共享 core 与 ui。
2. **Data Link Manager 单点职责**：引用 CRUD、解析、按源/按文档查询集中在一处，与事件驱动结合，便于维护。
3. **适配器模式**：Storage / File / Notification 的抽象与运行时注入，与「同源」需求匹配。
4. **Traceability 独立成服务**：版本、历史、依赖、审计集中，满足「完整数据溯源」。
5. **事件总线 + 插件**：Events 枚举与 EditorPlugin 接口为后续扩展（含 P2 协作、插件）留好接口。

---

## 6. 建议修改清单（按优先级）

| 优先级 | 项 | 操作 |
|--------|----|------|
| P0     | 统一 ReferenceType | 2.2 节 DataReference.type 与 3.2 一致，包含 `formula`。 |
| P0     | 统一 DataReference 定义 | 2.2 仅引用 3.2 的 DataSource/ReferenceTarget，不重复简略版。 |
| P0     | 图与模块对应关系 | 在 1.1 图或 2 节说明 Document Store Manager / Reference Service 与 Data Link Manager 的对应或分工。 |
| P1     | 选择性刷新 | SyncEngine 增加「按数据源/区域」刷新接口或说明。 |
| P1     | 冲突解决策略 | 2.3 或新小节简述 ResolveStrategy 种类与流程。 |
| P1     | Office 格式 | 增加「Office 格式兼容策略」小节（.docx/.xlsx 导入导出与 JSON 的关系）。 |
| P2     | 性能可验证性 | 6.1–6.3 补充性能目标与测试策略（2s/500ms/10k 行）。 |
| P2     | 协作与 50+ 用户 | 路线图或扩展架构中标注协作、多用户为 P2，与当前单机架构的关系。 |
| P2     | 开放问题 | 离线、大文件分块、协作冲突在架构中留扩展点或小节。 |

---

## 7. 代码实现与架构一致性检查

在完成 P0/P1 架构文档修订后，对当前代码与架构文档进行对照，结论如下。

### 7.1 已对齐项

| 项目 | 架构 | 代码 | 说明 |
|------|------|------|------|
| ReferenceType | 3.2 含 `cell \| range \| named \| formula` | `core/types.ts` 同 | 一致 |
| DataReference / DataSource / ReferenceTarget | 3.2 节定义 | `core/types.ts` 同（metadata 可选） | 一致 |
| DataLinkManager 核心 API | 2.2 create/get/update/delete, findBySource/ByDocument, resolve | `DataLinkManager.ts` 已实现 | 一致 |
| SyncEngine 刷新 | refreshReference, refreshDocument, refreshAll, **refreshBySource** | `DataSyncEngine.ts` 已全部实现 | P1 选择性刷新已落地 |
| ResolveStrategy | 2.3 冲突解决策略表 | `ConflictResolver.ts`: keep-local, keep-remote, latest-timestamp, merge, manual | 一致 |
| Events | 8.2 Events 枚举 | `eventBus.ts` 包含且扩展（ExcelWorkbookLoaded 等） | 一致 |
| TraceabilityService 能力 | 2.4 快照、历史、依赖、审计 | `TraceabilityService.ts` 已实现 | 一致 |

### 7.2 命名/签名差异（可保留，建议在架构中注明）

| 项目 | 架构 | 代码 | 建议 |
|------|------|------|------|
| DataLinkManager 事件 | onSourceChanged, onReferenceInvalid | onDataSourceChange, offDataSourceChange | 架构 2.2 可注明：实现中采用 onDataSourceChange（含订阅/取消），断链通过 resolve 失败时置 state='broken' 体现，无独立 onReferenceInvalid 回调。 |
| 解析返回值 | resolveReference(ref): any | resolveReference(ref): Promise\<any\> | 架构可注明为异步：resolveReference 返回 Promise\<any\>。 |
| Traceability 变更历史 | getChangeHistory(ref: DataReference): Change[] | getReferenceHistory(refId: ReferenceId): DataReference['history'] | 实现按 refId 查历史，与 DataReference.history 一致；架构可注明按 refId 的 getReferenceHistory。 |
| 快照创建 | createSnapshot 返回 SnapshotId | createSnapshot 返回 Promise\<SnapshotId\> | 实现为异步，架构可注明返回 Promise\<SnapshotId\>。 |

### 7.3 实现多于架构的合理扩展

以下为代码中已有、架构未写明的能力，属合理实现细节或扩展，无需改架构，仅列示便于对照：

- **DataLinkManager**：createReferenceWithId、setExcelWorkbook、getExcelDataSourceManager、getAllReferences、validateReference。
- **DataSyncEngine**：offDataChange、destroy、isAutoSyncEnabled。
- **TraceabilityService**：deleteSnapshot、clearAuditLog。
- **EventBus**：ExcelWorkbookLoaded、ExcelCellSelected、ExcelRangeSelected。

### 7.4 小结

- 当前实现与修订后的架构在**数据类型、核心 API、同步与冲突策略、溯源能力**上一致；P1 的 **refreshBySource** 已在代码与架构中同步补充。
- 差异主要为**事件命名与部分异步签名**，建议在架构 2.2、2.4 中做简短注明即可，无需改代码。
- 代码相对架构的扩展（Excel 集成、生命周期、审计清理）可视为实现细节，架构可在后续迭代中按需补充说明。

---

## 8. 总结

- **是否符合需求**：**是**。P0/P1 功能（引用类型、更新方式、溯源、Web/桌面同源、安全与扩展）在架构中均有对应，技术栈与需求一致。
- **是否合理**：**是**。分层、单一职责、适配器与事件驱动设计合理，可维护性和可扩展性良好。
- **建议**：优先完成「术语与模型统一」「图与正文对应」「选择性刷新与冲突策略」的补充；随后在兼容性、性能可测性、协作与开放问题上做轻量补充，即可作为实现与评审的稳定依据。
- **代码一致性**：P0/P1 修订已落实；实现与架构已对齐，仅存在少量命名/异步签名差异，已在 7 节列出并给出注明建议。

以上建议落实后，该架构文档可作为正式技术基准使用。
