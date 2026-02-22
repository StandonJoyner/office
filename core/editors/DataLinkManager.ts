import { eventBus, Events } from '../eventBus';
import {
  DataReference,
  DataSource,
  DocumentId,
  ReferenceId,
  ReferenceType,
} from '../types';
import type {
  IExcelDataSourceManager,
  IUniverWorkbook,
} from './ExcelDataSource';
import { createExcelDataSourceManager } from './ExcelDataSource';

export interface DataLinkManager {
  // Reference Management
  /** Creates a reference; target is the document id; stored target.nodeId is set to the new reference id. */
  createReference(source: DataSource, target: DocumentId): ReferenceId;
  /** Create a reference with a specific id (e.g. to match document node refId). */
  createReferenceWithId(id: ReferenceId, source: DataSource, target: DocumentId): ReferenceId;
  getReference(id: ReferenceId): DataReference | undefined;
  updateReference(id: ReferenceId, data: Partial<DataReference>): void;
  deleteReference(id: ReferenceId): void;

  // Reference Query
  /** Returns reference ids matching the given source; matching is by fileId and sheetId only. */
  findReferencesBySource(source: DataSource): ReferenceId[];
  findReferencesByDocument(docId: DocumentId): ReferenceId[];
  getAllReferences(): DataReference[];

  // Data Resolution
  /** Resolves value from data source and updates the reference's display and state; returns the value. */
  resolveReference(ref: DataReference): Promise<any>;
  resolveReferenceBatch(refs: DataReference[]): Promise<Map<ReferenceId, any>>;

  // Excel Integration
  setExcelWorkbook(workbook: IUniverWorkbook | null): void;
  getExcelDataSourceManager(): IExcelDataSourceManager;

  // Validation
  validateReference(ref: DataReference): boolean;

  // Data Source Change Monitoring
  onDataSourceChange(callback: (event: any) => void): void;
  offDataSourceChange(callback: (event: any) => void): void;

  // Lifecycle
  destroy(): void;
}

export class MemoryDataLinkManager implements DataLinkManager {
  private references = new Map<ReferenceId, DataReference>();
  private excelDataSourceManager: IExcelDataSourceManager;
  private dataSourceChangeCallbacks = new Set<(event: any) => void>();

  constructor(excelDataSourceManager?: IExcelDataSourceManager) {
    this.excelDataSourceManager = excelDataSourceManager ?? createExcelDataSourceManager();
    this.setupDataSourceMonitoring();
  }

  createReference(source: DataSource, target: DocumentId): ReferenceId {
    const id = crypto.randomUUID();
    return this.createReferenceWithId(id, source, target);
  }

  createReferenceWithId(id: ReferenceId, source: DataSource, target: DocumentId): ReferenceId {
    // Infer type from range: has endRow/endCol means range reference
    const type: ReferenceType = (source.range.endRow !== undefined && source.range.endCol !== undefined)
      ? 'range'
      : 'cell';

    const ref: DataReference = {
      id,
      type,
      source,
      target: { documentId: target, nodeId: id },
      display: { format: 'value' },
      state: 'active',
      history: [
        {
          timestamp: Date.now(),
          action: 'created',
          value: null,
          userId: 'system',
        },
      ],
    };
    this.references.set(id, ref);
    eventBus.emit(Events.ReferenceCreated, ref);
    return id;
  }

  getReference(id: ReferenceId): DataReference | undefined {
    return this.references.get(id);
  }

  updateReference(id: ReferenceId, data: Partial<DataReference>): void {
    const ref = this.references.get(id);
    if (ref) {
      const oldState = { ...ref };
      Object.assign(ref, data);
      eventBus.emit(Events.ReferenceUpdated, {
        referenceId: id,
        oldState,
        newState: ref,
      });
    }
  }

  deleteReference(id: ReferenceId): void {
    const ref = this.references.get(id);
    if (ref) {
      this.references.delete(id);
      eventBus.emit(Events.ReferenceDeleted, { referenceId: id });
    }
  }

  findReferencesBySource(source: DataSource): ReferenceId[] {
    const results: ReferenceId[] = [];
    this.references.forEach((ref, id) => {
      if (
        ref.source.fileId === source.fileId &&
        ref.source.sheetId === source.sheetId
      ) {
        results.push(id);
      }
    });
    return results;
  }

  findReferencesByDocument(docId: DocumentId): ReferenceId[] {
    const results: ReferenceId[] = [];
    this.references.forEach((ref, id) => {
      if (ref.target.documentId === docId) {
        results.push(id);
      }
    });
    return results;
  }

  getAllReferences(): DataReference[] {
    return Array.from(this.references.values());
  }

  /**
   * Set the Excel workbook for data source integration
   */
  setExcelWorkbook(workbook: IUniverWorkbook | null): void {
    if (workbook) {
      this.excelDataSourceManager.setWorkbook(workbook);
    }
  }

  /**
   * Get the Excel data source manager
   */
  getExcelDataSourceManager(): IExcelDataSourceManager {
    return this.excelDataSourceManager;
  }

  /**
   * Resolve a reference using the Excel data source manager
   */
  async resolveReference(ref: DataReference): Promise<any> {
    try {
      const value = await this.excelDataSourceManager.resolveReference(ref);

      // Update the reference with the resolved value
      this.updateReference(ref.id, {
        display: {
          ...ref.display,
          value,
        },
        state: 'active',
      });

      return value;
    } catch (error) {
      console.error(`Failed to resolve reference ${ref.id}:`, error);

      // Mark reference as broken
      this.updateReference(ref.id, { state: 'broken' });

      return ref.display.value; // Return cached value if available
    }
  }

  async resolveReferenceBatch(
    refs: DataReference[]
  ): Promise<Map<ReferenceId, any>> {
    const result = new Map<ReferenceId, any>();
    const promises = refs.map(async (ref) => {
      const value = await this.resolveReference(ref);
      return { id: ref.id, value };
    });

    const results = await Promise.all(promises);
    for (const { id, value } of results) {
      result.set(id, value);
    }
    return result;
  }

  /**
   * Validate that a reference's source exists
   */
  validateReference(ref: DataReference): boolean {
    if (ref.state === 'broken') return false;

    const { sheetId, range } = ref.source;
    return this.excelDataSourceManager.validateSheet(sheetId) &&
      this.excelDataSourceManager.validateCell(
        sheetId,
        range.startRow,
        range.startCol
      );
  }

  /**
   * Set up monitoring of Excel data source changes
   */
  private setupDataSourceMonitoring(): void {
    this.excelDataSourceManager.onDataSourceChange((event) => {
      // Find all references that match the changed cell
      const affectedRefs: DataReference[] = [];
      this.references.forEach((ref) => {
        if (ref.source.sheetId === event.sheetId) {
          if (ref.type === 'cell' ||
            (event.row !== undefined && event.column !== undefined &&
              ref.source.range.startRow === event.row &&
              ref.source.range.startCol === event.column)) {
            affectedRefs.push(ref);
          }
        }
      });

      // Update affected references
      affectedRefs.forEach(async (ref) => {
        const oldValue = ref.display.value;
        await this.resolveReference(ref);

        // Emit data change event
        eventBus.emit(Events.DataChanged, [{
          referenceId: ref.id,
          oldValue,
          newValue: ref.display.value,
          timestamp: Date.now(),
          source: 'auto',
        }]);
      });

      // Notify external listeners
      this.dataSourceChangeCallbacks.forEach((callback) => {
        try {
          callback(event);
        } catch (error) {
          console.error('Error in data source change callback:', error);
        }
      });
    });
  }

  /**
   * Register a data source change callback
   */
  onDataSourceChange(callback: (event: any) => void): void {
    this.dataSourceChangeCallbacks.add(callback);
  }

  /**
   * Unregister a data source change callback
   */
  offDataSourceChange(callback: (event: any) => void): void {
    this.dataSourceChangeCallbacks.delete(callback);
  }

  /**
   * Clean up resources
   */
  destroy(): void {
    this.excelDataSourceManager.destroy();
    this.dataSourceChangeCallbacks.clear();
    this.references.clear();
  }
}

/**
 * Factory function to create a DataLinkManager instance
 */
export function createDataLinkManager(
  excelDataSourceManager?: IExcelDataSourceManager
): DataLinkManager {
  return new MemoryDataLinkManager(excelDataSourceManager);
}
