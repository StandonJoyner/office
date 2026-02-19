# Office Suite PoC 验证计划

## 1. PoC 目标

验证核心技术在项目中的可行性，为完整开发提供技术依据。

### 1.1 验证重点
1. Tiptap 编辑器集成与扩展能力
2. Univer Excel 引擎功能覆盖度
3. Word 与 Excel 数据交互的技术方案
4. Web/桌面同源架构可行性
5. 性能基准测试

---

## 2. PoC 范围

### 2.1 功能范围
| 模块 | PoC 内容 | 完整开发范围 |
|------|----------|--------------|
| Word 编辑器 | 基础富文本 + 数据引用节点 | 完整文档编辑 + 协作 |
| Excel 编辑器 | 单表编辑 + 基础公式 | 多表 + 高级公式 + 图表 |
| 数据引用 | 单元引用 + 手动刷新 | 多种引用 + 自动同步 |
| 数据溯源 | 基础引用链 | 完整版本管理 |
| Web/桌面 | Web 版本验证 | 双平台发布 |

### 2.2 不包含的内容
- 协作编辑
- 云存储集成
- 插件系统
- 复杂的冲突解决
- 完整的权限管理

---

## 3. PoC 实施步骤

### Step 1: 项目脚手架 (1 天)

**目标**: 搭建可运行的开发环境

**任务**:
- [ ] 初始化 Next.js 项目（TypeScript, Tailwind）
- [ ] 配置开发工具（ESLint, Prettier）
- [ ] 集成测试框架（Vitest + Testing Library）
- [ ] 配置 CI/CD 基础

**验收标准**:
- 项目可启动，开发服务器正常运行
- 基础测试可执行
- 代码检查通过

---

### Step 2: Tiptap 编辑器集成 (2 天)

**目标**: 实现 Word 编辑器基础框架

**任务**:
- [ ] 安装并配置 Tiptap
- [ ] 实现基础富文本扩展（Bold, Italic, Heading, List）
- [ ] 创建数据引用节点 (ReferenceNode)
- [ ] 实现引用节点渲染组件
- [ ] 添加插入引用的命令

**代码示例**:
```typescript
// 定义引用节点
const ReferenceNode = Node.create({
  name: 'reference',
  group: 'inline',
  inline: true,
  atom: true,
  addAttributes() {
    return {
      refId: { default: null },
      value: { default: '' },
    };
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', { ...HTMLAttributes, class: 'data-reference' }];
  },
});
```

**验收标准**:
- 可进行基础文本编辑
- 可插入数据引用节点
- 引用节点正确显示

---

### Step 3: Univer Excel 引擎集成 (3 天)

**目标**: 实现 Excel 编辑器基础功能

**任务**:
- [ ] 安装并配置 Univer
- [ ] 实现基础表格渲染
- [ ] 实现单元格编辑
- [ ] 实现基础公式计算
- [ ] 添加单元格变更监听

**代码示例**:
```typescript
// 单元格变更监听
univerAPI.onCommandExecuted((command) => {
  if (command.id === SetRangeValuesCommand.id) {
    const changes = command.params.values;
    // 触发数据变更事件
    eventBus.emit(Events.DataChanged, { changes });
  }
});
```

**验收标准**:
- 可进行基础表格编辑
- 公式计算正确
- 可监听单元格变更

---

### Step 4: 数据链接管理器 (2 天)

**目标**: 实现 Word 与 Excel 之间的数据连接

**任务**:
- [ ] 定义数据模型（Reference, DataSource）
- [ ] 实现 DataLinkManager 类
- [ ] 实现引用创建与删除
- [ ] 实现引用值解析
- [ ] 实现引用有效性验证

**代码示例**:
```typescript
class DataLinkManager {
  private references = new Map<string, DataReference>();

  createReference(source: DataSource, docId: string): string {
    const id = uuid();
    const ref: DataReference = {
      id,
      type: 'cell',
      source,
      target: { documentId: docId },
      state: 'active',
    };
    this.references.set(id, ref);
    return id;
  }

  async resolveReference(ref: DataReference): Promise<any> {
    // 从 Univer 获取单元格值
    const workbook = univerAPI.getActiveWorkbook();
    const sheet = workbook.getSheetById(ref.source.sheetId);
    return sheet.getRange(ref.source.range).getValue();
  }
}
```

**验收标准**:
- 可创建和管理引用
- 可解析引用并获取数据
- 可检测无效引用

---

### Step 5: 数据引用与刷新 (2 天)

**目标**: 实现数据引用的展示和刷新功能

**任务**:
- [ ] 实现引用值显示
- [ ] 实现引用插入 UI（选择器）
- [ ] 实现手动刷新功能
- [ ] 实现引用跳转（点击跳转到数据源）
- [ ] 添加引用状态指示器

**代码示例**:
```typescript
// 刷新引用
async function refreshReference(refId: string) {
  const ref = dataLinkManager.getReference(refId);
  const value = await dataLinkManager.resolveReference(ref);

  // 更新引用节点
  const editor = wordEditor.getEditor();
  const { tr } = editor.state;
  const nodePos = findNodePosition(editor.state, refId);

  tr.setNodeMarkup(nodePos, null, {
    ...ref,
    value,
    state: 'active',
  });

  editor.view.dispatch(tr);
}
```

**验收标准**:
- 引用显示正确的数据值
- 可手动刷新引用
- 点击引用可跳转

---

### Step 6: 数据溯源（基础）(2 天)

**目标**: 实现基础的数据追溯能力

**任务**:
- [ ] 实现引用历史记录
- [ ] 实现依赖查询（反向依赖）
- [ ] 实现变更历史 UI
- [ ] 实现基础版本快照

**代码示例**:
```typescript
interface ReferenceHistory {
  timestamp: number;
  oldValue: any;
  newValue: any;
  action: 'created' | 'updated' | 'refreshed';
}

// 记录变更
function recordChange(refId: string, change: ReferenceHistory) {
  const ref = dataLinkManager.getReference(refId);
  ref.history.push(change);
}
```

**验收标准**:
- 可查看引用历史
- 可查看哪些文档引用了某单元格
- 可创建版本快照

---

### Step 7: Web/桌面适配验证 (1 天)

**目标**: 验证同源架构可行性

**任务**:
- [ ] 创建存储适配器接口
- [ ] 实现 Web 存储适配器（IndexedDB）
- [ ] 验证应用在浏览器中正常运行
- [ ] 评估桌面化技术选型（Tauri vs Electron）

**验收标准**:
- 应用在浏览器中正常运行
- 存储适配器设计合理
- 技术选型评估报告

---

### Step 8: 性能基准测试 (1 天)

**目标**: 评估核心性能指标

**测试项**:
- [ ] 文档打开时间（10页, 100页）
- [ ] 数据刷新延迟（10个引用, 100个引用）
- [ ] Excel 文件大小影响（100行, 1000行, 10000行）
- [ ] 内存占用

**验收标准**:
- 10页文档打开 < 1s
- 10个引用刷新 < 300ms
- 10000行 Excel 加载可接受（<5s）

---

## 4. 技术风险评估

| 风险项 | 风险等级 | 缓解措施 |
|--------|----------|----------|
| Tiptap 自定义节点复杂性 | 中 | 参考 ProseMirror 文档，从小范围开始 |
| Univer API 稳定性 | 中 | 使用稳定版本，必要时降级方案 |
| 数据同步冲突 | 高 | PoC 仅实现简单场景，后续完善 |
| 性能瓶颈 | 中 | 提前性能测试，必要时优化 |
| Web/桌面差异 | 低 | 设计良好的适配器层 |

---

## 5. 成功标准

### 5.1 功能标准
- [ ] Word 编辑器可进行基础编辑
- [ ] Excel 编辑器可编辑和计算
- [ ] 可在 Word 中插入数据引用
- [ ] 引用可刷新并显示正确值
- [ ] 可查看引用历史和依赖

### 5.2 技术标准
- [ ] 代码通过 ESLint 检查
- [ ] 核心功能有测试覆盖
- [ ] 架构设计符合规划
- [ ] 性能指标达标

### 5.3 文档标准
- [ ] 核心代码有注释
- [ ] 关键设计有文档
- [ ] 风险和限制有说明

---

## 6. 时间表

| 阶段 | 任务 | 预计时间 | 依赖 |
|------|------|----------|------|
| 1 | 项目脚手架 | 1 天 | 无 |
| 2 | Tiptap 集成 | 2 天 | 阶段 1 |
| 3 | Univer 集成 | 3 天 | 阶段 1 |
| 4 | 数据链接管理器 | 2 天 | 阶段 2, 3 |
| 5 | 数据引用与刷新 | 2 天 | 阶段 4 |
| 6 | 数据溯源（基础） | 2 天 | 阶段 5 |
| 7 | Web/桌面适配 | 1 天 | 阶段 6 |
| 8 | 性能测试 | 1 天 | 阶段 7 |
| 9 | 评估与总结 | 1 天 | 全部阶段 |

**总计**: 14 天

---

## 7. 交付物

1. **可运行的 PoC 应用**
   - Web 版本可部署访问
   - 完整功能演示

2. **源代码**
   - 清晰的项目结构
   - 代码注释
   - 单元测试

3. **技术评估报告**
   - 技术栈可行性结论
   - 性能基准数据
   - 风险与限制
   - 下一步建议

4. **演示文档**
   - 功能演示视频/截图
   - 使用说明
   - 代码示例

---

## 8. 后续决策点

### 8.1 Go/No-Go 决策
基于 PoC 结果，对以下技术点做出决策：
- [ ] Tiptap 是否适合本项目
- [ ] Univer 是否满足需求
- [ ] 数据同步方案是否可行
- [ ] Web/桌面同源架构是否继续

### 8.2 架构调整
根据 PoC 发现，可能需要调整：
- 数据模型设计
- 引用机制实现
- 性能优化策略

### 8.3 完整开发计划
基于 PoC 结果，细化完整开发计划：
- 功能优先级调整
- 资源分配
- 里程碑设置
