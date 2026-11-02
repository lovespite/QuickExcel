/** QuickExcel native bridge contract. No DOM, Node, filesystem, network or asynchronous host API. */
type QxErrorCode = 'Value' | 'Name' | 'Reference' | 'DivisionByZero' | 'Number' | 'NotAvailable' | 'Null' | 'Script' | 'Cycle' | 'Calculation';
type QxScalar = number | string | boolean | null | { readonly kind: 'error'; readonly code: QxErrorCode };
interface QxFunctionDescriptor {
  name: string;
  apiVersion: 2;
  parameters?: readonly { name: string; type?: 'any' | 'number' | 'text' | 'boolean' | 'blank' | 'range' }[];
  returns?: 'any' | 'number' | 'text' | 'boolean' | 'blank';
  profile?: 'conservative';
}
interface QxActionContext {
  readonly snapshot: {
    readonly baseRevision: number;
    readonly dateSystem: 'Windows1900' | 'Mac1904';
    readonly sheets: readonly { readonly id: string; readonly name: string }[];
    getCell(sheetId: string, cell: string): Readonly<{ value: QxScalar; formula: string | null; state: string; hasValue: boolean }>;
    getRange(sheetId: string, range: string): readonly (readonly QxScalar[])[];
  };
  readonly patch: {
    setValue(sheetId: string, cell: string, value: QxScalar): void;
    setValues(sheetId: string, range: string, values: readonly (readonly QxScalar[])[]): void;
    setFormula(sheetId: string, cell: string, formula: string): void;
    setFormulas(sheetId: string, range: string, formulas: readonly (readonly string[])[]): void;
  };
}
declare const quickExcel: {
  readonly functions: {
    /** Legacy scalar / flattened range arguments. Registration is allowed only during module initialization. */
    register(name: string, callback: (...args: any[]) => QxScalar): void;
    /** Descriptor API v2 preserves Blank and rectangular ranges; JS namespace required. */
    register(descriptor: QxFunctionDescriptor, callback: (...args: any[]) => QxScalar): void;
  };
  /** File scriptApiVersion 2 and an actions entry binding are required. */
  readonly actions: { register(name: string | { name: string; description?: string }, callback: (context: QxActionContext) => void | null): void };
};
declare const console: { log(...values: unknown[]): void; info(...values: unknown[]): void; warn(...values: unknown[]): void; error(...values: unknown[]): void };
