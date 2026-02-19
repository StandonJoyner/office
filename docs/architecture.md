# Office Suite 技术架构设计

## 1. 架构概览

### 1.1 整体架构图

```
┌─────────────────────────────────────────────────────────────────┐
│                        Presentation Layer                      │
├─────────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐             │
│  │   Word UI   │  │  Excel UI   │  │  Common UI  │             │
│  │  (React)    │  │  (React)    │  │  (React)    │             │
│  └──────┬──────┘  └──────┬──────┘  └──────┬──────┘             │
│         │                │                │                    │
│  ┌──────▼────────────────▼────────────────▼──────┐            │
│  │              Editor Core Layer                 │            │
│  │  ┌──────────┐      ┌──────────┐              │            │
│  │  │ Tiptap   │      │  Univer  │              │            │
│  │  │  Core    │      │  Core    │              │            │
│  │  └────┬─────┘      └────┬─────┘              │            │
│  │       │                 │                     │            │
│  │       └────────┬────────┘                     │            │
│  │                ▼                              │            │
│  │        ┌──────────────┐                      │            │
│  │        │ Data Link    │                      │            │
│  │        │    Manager   │                      │            │
│  │        └──────┬───────┘                      │            │
│  └───────────────┼───────────────────────────────┘            │
└──────────────────┼───────────────────────────────────────────────┘
                   │
┌──────────────────▼───────────────────────────────────────────────┐
│                       Business Logic Layer                       │
├─────────────────────────────────────────────────────────────────┤
│  ┌──────────────────┐  ┌──────────────────┐                     │
│  │  Document Store  │  │   Reference      │                     │
│  │    Manager       │  │    Service       │                     │
│  └────────┬─────────┘  └────────┬─────────┘                     │
│           │                      │                               │
│  ┌────────▼──────────┐  ┌────────▼──────────┐                     │
│  │  Excel Data       │  │  Sync Engine      │                     │
│  │   Manager         │  │                   │                     │
│  └────────┬──────────┘  └────────┬──────────┘                     │
│           │                      │                               │
│  ┌────────▼──────────────────────▼──────────┐                    │
│  │          Data Traceability Service       │                    │
│  │  - Version tracking                      │                    │
│  │  - Change history                         │                    │
│  │  - Impact analysis                        │                    │
│  └──────────────────────┬───────────────────┘                    │
└─────────────────────────┼───────────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────────┐
│                        Data Layer                                │
├─────────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐ │
│  │  In-Memory      │  │  IndexedDB      │  │  File System    │ │
│  │  Store          │  │  (Browser)      │  │  (Desktop)      │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘ │
│                                                                   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │               Data Serialization Format                   │   │
│  │  - Document JSON (Tiptap/Univer compatible)               │   │
│  │  - Reference Metadata                                     │   │
│  │  - Version Snapshots                                     │   │
│  └──────────────────────────────────────────────────────────┘   │
└───────────────────────────────────────────────────────────────────┘

                          ┌─────────────────┐
                          │  Runtime Layer   │
                          ├─────────────────┤
                          │  Browser (Web)  │
                          │  Electron       │
                          │  Tauri          │
                          └─────────────────┘
```

**图例与模块对应**：业务逻辑层中的 **Reference Service** 即 **Data Link Manager**（数据链接管理器），负责引用的创建、解析与按源/按文档查询；**Document Store Manager** 负责文档与工作簿的持久化、元数据及与 Data Layer 的交互。Editor Core 中的 Data Link Manager 与业务层的 Reference Service 为同一抽象在不同层级的体现（核心能力在 Editor Core，持久化与协同在 Business Logic）。

### 1.2 核心设计原则

| 原则 | 说明 |
|------|------|
| 同构设计 | Web 和桌面共享同一套业务逻辑 |
| 模块化 | 编辑器核心与业务逻辑解耦 |
| 事件驱动 | 通过事件总线进行跨模块通信 |
| 可扩展 | 插件架构支持功能扩展 |
| 类型安全 | 全面使用 TypeScript |

---

## 2. 核心模块设计

### 2.1 编辑器核心层 (Editor Core Layer)

#### 2.1.1 Word 编辑器 (Tiptap)

```
┌─────────────────────────────────────────────┐
│           WordEditor Component              │
├─────────────────────────────────────────────┤
│  Tiptap Editor Instance                     │
│  ├─ Extensions                              │
│  │   ├─ Rich Text (Bold, Italic, ...)      │
│  │   ├─ Tables                             │
│  │   ├─ Images                             │
│  │   └─ DataReference (自定义)             │
│  ├─ Node System                            │
│  │   └─ ReferenceNode (数据引用节点)        │
│  └─ Mark System                            │
│      └─ ReferenceMark (行内引用标记)        │
└─────────────────────────────────────────────┘
```

**DataReference 扩展**
```typescript
interface DataReferenceExtension {
  // 引用类型定义（与 3.2 ReferenceType 一致）
  referenceTypes: {
    cell: CellReference;      // 单元格引用
    range: RangeReference;    // 范围引用
    named: NamedReference;    // 命名区域引用
    formula: FormulaReference; // 公式引用（公式计算结果）
  };

  // 插入引用
  insertReference(ref: DataReference): void;

  // 更新引用值
  updateReferenceValue(id: string, value: any): void;

  // 获取所有引用
  getAllReferences(): DataReference[];

  // 验证引用有效性
  validateReference(ref: DataReference): boolean;
}
```

#### 2.1.2 Excel 引擎 (Univer)

```
┌─────────────────────────────────────────────┐
│          ExcelEditor Component              │
├─────────────────────────────────────────────┤
│  Univer Instance                            │
│  ├─ Workbook                                │
│  │   ├─ Worksheets                          │
│  │   │   └─ Cells (Data + Formula)         │
│  │   └─ Named Ranges                        │
│  ├─ Calculation Engine                      │
│  └─ Custom API (扩展)                       │
│      ├─ CellChangeTracker                   │
│      ├─ DependencyAnalyzer                  │
│      └─ ValueSnapshot                      │
└─────────────────────────────────────────────┘
```

### 2.2 数据链接管理器 (Data Link Manager)

这是连接 Word 和 Excel 的核心模块。

```typescript
interface DataLinkManager {
  // 引用管理
  createReference(source: DataSource, target: DocumentId): ReferenceId;
  getReference(id: ReferenceId): DataReference;
  updateReference(id: ReferenceId, data: any): void;
  deleteReference(id: ReferenceId): void;

  // 引用查询
  findReferencesBySource(source: DataSource): ReferenceId[];
  findReferencesByDocument(docId: DocumentId): ReferenceId[];

  // 数据获取
  resolveReference(ref: DataReference): any;
  resolveReferenceBatch(refs: DataReference[]): Map<ReferenceId, any>;

  // 事件监听
  onSourceChanged(callback: (source: DataSource) => void): void;
  onReferenceInvalid(callback: (ref: DataReference) => void): void;
}

// DataReference 完整定义见 3.2 节（ReferenceType = 'cell' | 'range' | 'named' | 'formula'，
// 以及 DataSource、ReferenceTarget、ReferenceDisplay、ReferenceState、ReferenceHistory）
```

### 2.3 同步引擎 (Sync Engine)

```typescript
interface SyncEngine {
  // 手动刷新
  refreshReference(id: ReferenceId): Promise<void>;
  refreshDocument(docId: DocumentId): Promise<void>;
  refreshAll(): Promise<void>;
  /** 按数据源刷新：刷新来自指定 DataSource 的所有引用（选择性刷新） */
  refreshBySource(source: DataSource): Promise<void>;

  // 自动同步
  enableAutoSync(enabled: boolean): void;
  setSyncInterval(interval: number): void;

  // 冲突检测
  detectConflicts(): Conflict[];
  resolveConflict(conflict: Conflict, strategy: ResolveStrategy): void;

  // 变更通知
  onDataChange(callback: (changes: DataChange[]) => void): void;
}

interface DataChange {
  referenceId: string;
  oldValue: any;
  newValue: any;
  timestamp: number;
  source: 'manual' | 'auto';
}
```

#### 冲突解决策略 (ResolveStrategy)

| 策略 | 说明 |
|------|------|
| `keep-local` | 以当前文档/本地值为准，丢弃远端变更 |
| `keep-remote` | 以数据源/远端值为准，覆盖本地 |
| `latest-timestamp` | 以时间戳较新的一方为准 |
| `merge` | 尝试合并（若可合并）；否则标记需人工处理 |
| `manual` | 不自动决定，由用户在选择界面中指定最终值 |

流程：`detectConflicts()` 返回冲突列表 → 每项可带 `suggestedStrategy` → 调用 `resolveConflict(conflict, strategy)` 应用策略并更新引用值、标记已解决。

### 2.4 数据溯源服务 (Data Traceability Service)

```typescript
interface TraceabilityService {
  // 版本追踪
  createSnapshot(docId: DocumentId): SnapshotId;
  getSnapshot(id: SnapshotId): DocumentSnapshot;
  listSnapshots(docId: DocumentId): DocumentSnapshot[];
  compareSnapshots(id1: SnapshotId, id2: SnapshotId): Diff[];

  // 变更历史
  getChangeHistory(ref: DataReference): Change[];
  getCellHistory(cellId: CellId): CellChange[];

  // 影响分析
  getDependents(cellId: CellId): ReferenceId[];
  getDependencies(docId: DocumentId): DataSource[];

  // 审计日志
  logAction(action: AuditAction): void;
  getAuditLog(filters: AuditFilters): AuditLog[];
}

interface DocumentSnapshot {
  id: string;
  documentId: string;
  timestamp: number;
  data: {
    content: any; // Tiptap/Univer JSON
    references: DataReference[];
  };
  checksum: string;
}
```

---

## 3. 数据模型设计

### 3.1 文档数据模型

```typescript
// Word 文档
interface WordDocument {
  id: string;
  name: string;
  content: TiptapJSON; // Tiptap 编辑器 JSON
  references: DataReference[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}

// Excel 文档
interface ExcelDocument {
  id: string;
  name: string;
  workbook: UniverWorkbookJSON;
  namedRanges: NamedRange[];
  metadata: DocumentMetadata;
  version: number;
  createdAt: number;
  updatedAt: number;
}

interface DocumentMetadata {
  author: string;
  tags: string[];
  isTemplate: boolean;
  parentId?: string; // 模板来源
}
```

### 3.2 引用数据模型

```typescript
interface DataReference {
  id: string; // UUID
  type: ReferenceType;
  source: DataSource;
  target: ReferenceTarget;
  display: ReferenceDisplay;
  state: ReferenceState;
  history: ReferenceHistory[];
}

type ReferenceType = 'cell' | 'range' | 'named' | 'formula';

interface DataSource {
  fileId: string;
  fileName: string;
  sheetId: string;
  sheetName: string;
  range: CellRange;
  isFormula: boolean;
}

interface CellRange {
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
}

interface ReferenceTarget {
  documentId: string;
  nodeId: string; // Tiptap node ID
  offset?: number;
}

interface ReferenceDisplay {
  format: 'value' | 'expression' | 'mixed';
  value?: any;
  expression?: string;
  tooltip?: string;
}

type ReferenceState = 'active' | 'stale' | 'broken' | 'conflict';

interface ReferenceHistory {
  timestamp: number;
  action: 'created' | 'updated' | 'resolved' | 'broken';
  value: any;
  userId: string;
}
```

### 3.3 存储格式

```
document.office
├── meta.json          # 文档元数据
├── content.json       # 编辑器内容 (Tiptap/Univer JSON)
├── references.json    # 引用列表
├── snapshots/         # 版本快照
│   ├── v1.json
│   └── v2.json
└── history.json       # 变更历史
```

### 3.4 Office 格式兼容策略

| 方面 | 策略 |
|------|------|
| **内部表示** | 编辑与存储统一使用 Tiptap JSON（Word）、Univer Workbook JSON（Excel）及 3.3 的 document.office 结构；引用、版本、历史均基于该模型。 |
| **导入** | .docx/.xlsx 通过转换层解析为上述 JSON 后写入 content/workbook；转换层可置于 Data Layer 或独立服务，优先级可放在 P1（文件系统支持）或 P2。 |
| **导出** | 由 Tiptap/Univer JSON 生成 .docx/.xlsx，保证与「单一内部格式」一致；导出时引用可渲染为当前值或保留为域/链接。 |
| **兼容性目标** | 满足需求 5.2：支持打开、编辑、保存为 Office 格式；具体兼容粒度（样式、复杂对象）在实现阶段定义。 |

---

## 4. Web/桌面同源架构

### 4.1 架构对比

| 层级 | Web | Desktop | 共享代码 |
|------|-----|---------|----------|
| UI 组件 | React DOM | React DOM | 100% |
| 编辑器核心 | Tiptap/Univer | Tiptap/Univer | 100% |
| 业务逻辑 | JavaScript | JavaScript | 100% |
| 状态管理 | React Query | React Query | 100% |
| 文件系统 | File API | Node.js fs | 适配层 |
| 存储 | IndexedDB | 文件系统 | 存储适配器 |
| 通知 | Service Worker | Electron IPC | 通知适配器 |

### 4.2 适配器模式

```typescript
// 抽象接口
interface StorageAdapter {
  save(key: string, value: any): Promise<void>;
  load(key: string): Promise<any>;
  delete(key: string): Promise<void>;
}

// Web 实现
class BrowserStorageAdapter implements StorageAdapter {
  async save(key: string, value: any): Promise<void> {
    await idb.set(key, value);
  }
  // ...
}

// Desktop 实现
class DesktopStorageAdapter implements StorageAdapter {
  async save(key: string, value: any): Promise<void> {
    await window.api.storage.save(key, value);
  }
  // ...
}

// 运行时注入
const storage: StorageAdapter = isDesktop
  ? new DesktopStorageAdapter()
  : new BrowserStorageAdapter();
```

### 4.3 构建策略

项目采用根目录分层（无 `src/` 前缀），与 Next.js App Router 一致：

```
├── core/           # 核心业务逻辑（平台无关）
│   ├── editors/    # 编辑器抽象、DataLinkManager、Excel 数据源
│   ├── sync/       # 同步引擎、冲突解决
│   └── trace/      # 溯源服务
├── ui/             # UI 组件（平台无关）
│   ├── word/       # Word 编辑器组件
│   ├── excel/      # Excel 编辑器组件（Univer）
│   ├── common/     # 通用组件（TraceabilityPanel、ReferenceDetailsPanel 等）
│   ├── components/ # 业务组件（DataUpdateNotification、DataUpdateConfirmation）
│   └── hooks/      # 与 core 集成的 React hooks
├── adapters/       # 平台适配器
│   ├── web/        # Web 实现（File/System）
│   └── tauri/      # Desktop 实现（Tauri）
└── app/            # Next.js 应用入口（App Router）
    ├── layout.tsx
    ├── page.tsx
    ├── integrated/ # Word+Excel 集成页
    └── excel/      # Excel 单页
```

---

## 5. 技术栈详细说明

### 5.1 前端框架

| 技术 | 版本 | 用途 |
|------|------|------|
| React | 18.3+ | UI 框架 |
| Next.js | 14+ (App Router) | Web 框架 |
| TypeScript | 5.0+ | 类型系统 |
| Tailwind CSS | 3.4+ | 样式 |

### 5.2 编辑器引擎

| 技术 | 版本 | 用途 |
|------|------|------|
| Tiptap | 2.0+ | Word 编辑器 |
| ProseMirror | 1.0+ | Tiptap 底层 |
| Univer | 0.1+ | Excel 引擎 |

### 5.3 状态管理

| 技术 | 版本 | 用途 |
|------|------|------|
| Zustand | 4.4+ | 应用状态 |
| React Query | 5.0+ | 服务端状态 |
| Immer | 10.0+ | 不可变更新 |

### 5.4 桌面化方案

| 方案 | 优点 | 缺点 | 选择 |
|------|------|------|------|
| Electron | 生态成熟，社区大 | 体积大 | 备选 |
| Tauri | 体积小，性能好 | 生态较新 | 推荐 |

---

## 6. 性能优化策略

### 6.1 编辑器性能

- **虚拟滚动**：长文档渲染优化
- **延迟加载**：图片、附件按需加载
- **增量更新**：只更新变化的内容节点

### 6.2 数据引用性能

- **缓存策略**：引用值缓存，减少重复计算
- **批量刷新**：合并多个引用的更新请求
- **异步解析**：大型数据源异步加载

### 6.3 文件操作性能

- **流式读写**：大文件分块处理
- **索引缓存**：常用文件内容缓存
- **预取策略**：预测用户行为，提前加载

---

## 7. 安全架构

### 7.1 数据安全
- 内容加密：AES-256
- 传输加密：TLS 1.3
- 敏感数据脱敏

### 7.2 访问控制
- 文档级权限
- 引用级权限
- 操作审计日志

---

## 8. 扩展架构

### 8.1 插件系统

```typescript
interface EditorPlugin {
  name: string;
  version: string;
  init(editor: Editor): void;
  destroy(): void;
  extensions?: Extension[];
  commands?: Commands;
}

// 示例：数据可视化插件
class ChartPlugin implements EditorPlugin {
  name = 'chart';
  init(editor: Editor) {
    // 注册图表节点
    editor.registerNode(ChartNode);
  }
}
```

### 8.2 事件总线

```typescript
interface EventBus {
  on(event: string, handler: Handler): void;
  off(event: string, handler: Handler): void;
  emit(event: string, payload: any): void;
}

// 事件类型
enum Events {
  ReferenceCreated = 'reference:created',
  ReferenceUpdated = 'reference:updated',
  ReferenceBroken = 'reference:broken',
  DataChanged = 'data:changed',
  DocumentSaved = 'document:saved',
}
```

---

## 9. 部署架构

### 9.1 Web 部署

```
┌─────────────┐
│   CDN/Edge  │  静态资源分发
└──────┬──────┘
       │
┌──────▼──────┐
│   Next.js   │  SSR/SSG
│   App       │
└──────┬──────┘
       │
┌──────▼──────┐
│  Node.js    │  后端服务
└─────────────┘
```

### 9.2 Desktop 打包

- Tauri：Rust + Web 前端
- 打包格式：.exe, .dmg, .AppImage

---

## 10. 开发路线图

### Phase 1: 基础架构 (Weeks 1-2)
- [x] 项目脚手架搭建
- [ ] Tiptap 基础集成
- [ ] Univer 基础集成
- [ ] 数据模型定义

### Phase 2: 核心功能 (Weeks 3-6)
- [ ] 数据引用系统
- [ ] 数据链接管理器
- [ ] 基础同步机制
- [ ] 数据溯源（基础）

### Phase 3: 增强功能 (Weeks 7-10)
- [ ] 自动同步
- [ ] 冲突检测
- [ ] 版本管理
- [ ] 文件系统集成

### Phase 4: 优化与桌面化 (Weeks 11-14)
- [ ] 性能优化
- [ ] 桌面适配
- [ ] 打包部署
- [ ] 文档完善
