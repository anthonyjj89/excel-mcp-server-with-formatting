declare module '@modelcontextprotocol/sdk' {
  export class FastMCP {
    constructor(options?: any);
    addTool(tool: any): void;
    setErrorHandler(handler: (error: any) => any): void;
    run(): Promise<void>;
    run_sse_async(port?: number): Promise<void>;
  }

  export class UserError extends Error {
    constructor(message: string);
  }

  export class Tool {
    constructor(name: string, options?: any);
    execute(params: any): Promise<any>;
  }

  export function createServer(options: any): any;
  export function createTool(tool: any): any;
  export function createStdioTransport(): any;
}
