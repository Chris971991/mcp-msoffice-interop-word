/**
 * Debug utility for MCP server logging
 * Only logs to console when not in stdio mode to avoid interfering with JSON protocol
 */

class DebugLogger {
  private isStdioMode: boolean;

  constructor() {
    this.isStdioMode = (process.env.MCP_TRANSPORT || 'stdio') === 'stdio';
  }

  /**
   * Log informational messages (only in non-stdio mode)
   */
  log(message: string, ...args: any[]): void {
    if (!this.isStdioMode) {
      console.log(message, ...args);
    }
  }

  /**
   * Log warning messages (only in non-stdio mode)
   */
  warn(message: string, ...args: any[]): void {
    if (!this.isStdioMode) {
      console.warn(message, ...args);
    }
  }

  /**
   * Log error messages (only in non-stdio mode)
   * In stdio mode, errors should be handled via MCP error responses
   */
  error(message: string, ...args: any[]): void {
    if (!this.isStdioMode) {
      console.error(message, ...args);
    }
  }

  /**
   * Always log critical errors regardless of mode (to stderr)
   */
  critical(message: string, ...args: any[]): void {
    console.error(message, ...args);
  }
}

// Export singleton instance
export const debug = new DebugLogger();