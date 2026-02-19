/**
 * Performance Monitor for Data Synchronization
 *
 * Tracks and analyzes sync performance metrics
 */

export interface SyncMetrics {
  operationId: string;
  operation: string;
  startTime: number;
  endTime?: number;
  duration?: number;
  referenceCount: number;
  success: boolean;
  error?: string;
}

export interface PerformanceStats {
  totalOperations: number;
  successfulOperations: number;
  failedOperations: number;
  averageDuration: number;
  minDuration: number;
  maxDuration: number;
  p95Duration: number;
  p99Duration: number;
  operationsPerSecond: number;
}

export interface PerformanceThresholds {
  warningThreshold: number; // ms
  errorThreshold: number; // ms
  alertThreshold: number; // ops/min
}

const DEFAULT_THRESHOLDS: PerformanceThresholds = {
  warningThreshold: 500, // 500ms
  errorThreshold: 1000, // 1s
  alertThreshold: 60, // 60 ops/min
};

export interface PerformanceAlert {
  id: string;
  type: 'slow-operation' | 'high-error-rate' | 'degraded-performance';
  severity: 'info' | 'warning' | 'error' | 'critical';
  message: string;
  timestamp: number;
  metrics?: any;
}

/**
 * Performance Monitor
 *
 * Tracks sync operations and provides performance analytics
 */
export class PerformanceMonitor {
  private readonly metrics: SyncMetrics[] = [];
  private readonly alertCallbacks = new Set<(alert: PerformanceAlert) => void>();
  private readonly thresholds: PerformanceThresholds;
  private operationCount = 0;
  private windowStartTime = Date.now();

  constructor(thresholds?: Partial<PerformanceThresholds>) {
    this.thresholds = { ...DEFAULT_THRESHOLDS, ...thresholds };
  }

  /**
   * Start tracking an operation
   */
  startOperation(operation: string, referenceCount: number = 1): string {
    const operationId = `${operation}-${Date.now()}-${crypto.randomUUID()}`;

    this.metrics.push({
      operationId,
      operation,
      startTime: Date.now(),
      referenceCount,
      success: false,
    });

    return operationId;
  }

  /**
   * End tracking an operation
   */
  endOperation(operationId: string, success: boolean = true, error?: string): void {
    const metric = this.metrics.find(m => m.operationId === operationId);

    if (!metric) {
      console.warn(`Operation ${operationId} not found`);
      return;
    }

    metric.endTime = Date.now();
    metric.duration = metric.endTime - metric.startTime;
    metric.success = success;

    if (error) {
      metric.error = error;
    }

    this.operationCount++;

    // Check for slow operations
    if (metric.success && metric.duration! > this.thresholds.warningThreshold) {
      this.emitAlert({
        id: `slow-${operationId}`,
        type: 'slow-operation',
        severity: 'warning',
        message: `Operation '${metric.operation}' took ${metric.duration}ms (threshold: ${this.thresholds.warningThreshold}ms)`,
        timestamp: Date.now(),
        metrics: { duration: metric.duration },
      });
    }

    // Check for errors
    if (!metric.success) {
      this.emitAlert({
        id: `error-${operationId}`,
        type: 'slow-operation',
        severity: 'error',
        message: `Operation '${metric.operation}' failed: ${error}`,
        timestamp: Date.now(),
        metrics: { error },
      });
    }
  }

  /**
   * Get statistics for all operations
   */
  getStats(timeWindowMs: number = 60000): PerformanceStats {
    const now = Date.now();
    const recentMetrics = this.metrics.filter(
      m => m.startTime >= now - timeWindowMs
    );

    if (recentMetrics.length === 0) {
      return {
        totalOperations: 0,
        successfulOperations: 0,
        failedOperations: 0,
        averageDuration: 0,
        minDuration: 0,
        maxDuration: 0,
        p95Duration: 0,
        p99Duration: 0,
        operationsPerSecond: 0,
      };
    }

    const durations = recentMetrics
      .filter(m => m.duration !== undefined)
      .map(m => m.duration!) as number[];

    const sortedDurations = [...durations].sort((a, b) => a - b);

    const successful = recentMetrics.filter(m => m.success);
    const failed = recentMetrics.filter(m => !m.success);

    // Calculate throughput
    const windowDuration = now - Math.min(...recentMetrics.map(m => m.startTime));
    const operationsPerSecond = windowDuration > 0
      ? Math.round((recentMetrics.length / windowDuration) * 1000)
      : 0;

    return {
      totalOperations: recentMetrics.length,
      successfulOperations: successful.length,
      failedOperations: failed.length,
      averageDuration: this.average(durations),
      minDuration: Math.min(...durations),
      maxDuration: Math.max(...durations),
      p95Duration: this.percentile(sortedDurations, 95),
      p99Duration: this.percentile(sortedDurations, 99),
      operationsPerSecond,
    };
  }

  /**
   * Get statistics by operation type
   */
  getStatsByOperation(operation: string, timeWindowMs: number = 60000): PerformanceStats {
    const now = Date.now();
    const recentMetrics = this.metrics.filter(
      m => m.operation === operation && m.startTime >= now - timeWindowMs
    );

    if (recentMetrics.length === 0) {
      return {
        totalOperations: 0,
        successfulOperations: 0,
        failedOperations: 0,
        averageDuration: 0,
        minDuration: 0,
        maxDuration: 0,
        p95Duration: 0,
        p99Duration: 0,
        operationsPerSecond: 0,
      };
    }

    const durations = recentMetrics
      .filter(m => m.duration !== undefined)
      .map(m => m.duration!) as number[];

    const sortedDurations = [...durations].sort((a, b) => a - b);

    const successful = recentMetrics.filter(m => m.success);
    const failed = recentMetrics.filter(m => !m.success);

    const windowDuration = now - Math.min(...recentMetrics.map(m => m.startTime));
    const operationsPerSecond = windowDuration > 0
      ? Math.round((recentMetrics.length / windowDuration) * 1000)
      : 0;

    return {
      totalOperations: recentMetrics.length,
      successfulOperations: successful.length,
      failedOperations: failed.length,
      averageDuration: this.average(durations),
      minDuration: durations.length > 0 ? Math.min(...durations) : 0,
      maxDuration: durations.length > 0 ? Math.max(...durations) : 0,
      p95Duration: this.percentile(sortedDurations, 95),
      p99Duration: this.percentile(sortedDurations, 99),
      operationsPerSecond,
    };
  }

  /**
   * Calculate average
   */
  private average(values: number[]): number {
    if (values.length === 0) return 0;
    return values.reduce((sum, v) => sum + v, 0) / values.length;
  }

  /**
   * Calculate percentile
   */
  private percentile(sortedValues: number[], p: number): number {
    if (sortedValues.length === 0) return 0;

    const index = Math.ceil((p / 100) * sortedValues.length) - 1;
    return sortedValues[Math.max(0, index)];
  }

  /**
   * Check if performance is degraded
   */
  isPerformanceDegraded(timeWindowMs: number = 60000): boolean {
    const stats = this.getStats(timeWindowMs);

    // High error rate (more than 10%)
    if (stats.totalOperations > 10) {
      const errorRate = stats.failedOperations / stats.totalOperations;
      if (errorRate > 0.1) {
        return true;
      }
    }

    // Low throughput
    if (stats.operationsPerSecond < 10) {
      return true;
    }

    // High latency
    if (stats.p95Duration > this.thresholds.errorThreshold) {
      return true;
    }

    return false;
  }

  /**
   * Register a callback for performance alerts
   */
  onAlert(callback: (alert: PerformanceAlert) => void): () => void {
    this.alertCallbacks.add(callback);

    return () => {
      this.alertCallbacks.delete(callback);
    };
  }

  /**
   * Emit an alert to all listeners
   */
  private emitAlert(alert: PerformanceAlert): void {
    this.alertCallbacks.forEach(callback => {
      try {
        callback(alert);
      } catch (error) {
        console.error('Error in performance alert callback:', error);
      }
    });
  }

  /**
   * Get recent metrics
   */
  getRecentMetrics(limit: number = 100): SyncMetrics[] {
    return this.metrics.slice(-limit);
  }

  /**
   * Get metrics by operation ID
   */
  getMetric(operationId: string): SyncMetrics | undefined {
    return this.metrics.find(m => m.operationId === operationId);
  }

  /**
   * Clear old metrics
   */
  clearOldMetrics(maxAge: number = 3600000): number { // 1 hour default
    const now = Date.now();
    const beforeCount = this.metrics.length;

    // Remove metrics older than maxAge
    for (let i = this.metrics.length - 1; i >= 0; i--) {
      if (now - this.metrics[i].startTime! > maxAge) {
        this.metrics.splice(i, 1);
      }
    }

    return beforeCount - this.metrics.length;
  }

  /**
   * Clear all metrics
   */
  clear(): void {
    this.metrics.length = 0;
    this.operationCount = 0;
    this.windowStartTime = Date.now();
  }

  /**
   * Get a performance report
   */
  getReport(timeWindowMs: number = 60000): string {
    const stats = this.getStats(timeWindowMs);
    const byOperation = new Map<string, PerformanceStats>();

    const operations = [...new Set(this.metrics.map(m => m.operation))];
    for (const op of operations) {
      byOperation.set(op, this.getStatsByOperation(op, timeWindowMs));
    }

    let report = `=== Performance Report (${timeWindowMs}ms window) ===\n\n`;
    report += `Total Operations: ${stats.totalOperations}\n`;
    report += `Successful: ${stats.successfulOperations} (${((stats.successfulOperations / stats.totalOperations) * 100).toFixed(1)}%)\n`;
    report += `Failed: ${stats.failedOperations} (${((stats.failedOperations / stats.totalOperations) * 100).toFixed(1)}%)\n`;
    report += `Average Duration: ${stats.averageDuration.toFixed(0)}ms\n`;
    report += `Min Duration: ${stats.minDuration}ms\n`;
    report += `Max Duration: ${stats.maxDuration}ms\n`;
    report += `P95 Duration: ${stats.p95Duration}ms\n`;
    report += `P99 Duration: ${stats.p99Duration}ms\n`;
    report += `Throughput: ${stats.operationsPerSecond} ops/sec\n`;

    report += '\n=== By Operation ===\n';
    for (const [op, opStats] of byOperation.entries()) {
      report += `\n${op}:\n`;
      report += `  Count: ${opStats.totalOperations}\n`;
      report += `  Avg Duration: ${opStats.averageDuration.toFixed(0)}ms\n`;
      report += `  P95: ${opStats.p95Duration}ms\n`;
    }

    return report;
  }

  /**
   * Export metrics as JSON
   */
  exportMetrics(): string {
    return JSON.stringify({
      stats: this.getStats(),
      recentMetrics: this.getRecentMetrics(1000),
      alerts: this.getAlerts(),
    }, null, 2);
  }

  /**
   * Get recent alerts
   */
  getAlerts(limit: number = 50): PerformanceAlert[] {
    // This would need alert tracking, simplified here
    return [];
  }

  /**
   * Destroy the monitor
   */
  destroy(): void {
    this.clear();
    this.alertCallbacks.clear();
  }
}

export function createPerformanceMonitor(thresholds?: Partial<PerformanceThresholds>): PerformanceMonitor {
  return new PerformanceMonitor(thresholds);
}
