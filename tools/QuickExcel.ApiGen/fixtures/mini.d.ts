// Synthetic fixture shaped like the Office Scripts definition. Used to verify the
// generator without downloading the pinned source.
declare namespace ExcelScript {
  /** Represents a color. */
  enum RangeValueType {
    double = "Double",
    string = "String",
    boolean = "Boolean",
    error = "Error",
    empty = "Empty"
  }

  /**
   * Represents the workbook.
   */
  interface Workbook {
    /** Gets the active cell. */
    getActiveCell(): Range;
    getActiveWorksheet(): Worksheet;
    getWorksheet(key: string): Worksheet;
    getWorksheets(): Worksheet[];
    getSelectedRange(): Range;
    getSelectedRanges(): RangeAreas;
    /** Not supported by QuickExcel yet. */
    getPivotTables(): PivotTable[];
    getApplication(): Application;
  }

  interface Worksheet {
    getId(): string;
    getName(): string;
    getCell(row: number, column: number): Range;
    getRange(address?: string): Range;
    getTables(): Table[];
  }

  interface Range {
    getAddress(): string;
    getWorksheet(): Worksheet;
    getValue(): string | number | boolean;
    getValues(): (string | number | boolean)[][];
    setValue(value: string | number | boolean): void;
    setValues(values: (string | number | boolean)[][]): void;
    getFormula(): string;
    setFormula(formula: string): void;
    setFormulas(formulas: string[][]): void;
    getFormat(): RangeFormat;
    getValueTypes(): RangeValueType[][];
    clear(applyTo?: string): void;
  }

  interface RangeAreas {
    getAddress(): string;
    getAreaCount(): number;
    getCellCount(): number;
    getAreas(): Range[];
    getWorksheet(): Worksheet;
    clear(applyTo?: string): void;
  }

  interface RangeFormat {
    getFill(): RangeFill;
  }

  interface RangeFill {
    getColor(): string;
    setColor(color: string): void;
    clear(): void;
  }

  interface PivotTable {
    getName(): string;
  }

  interface Table {
    getName(): string;
  }

  interface Application {
    calculate(calculationType: CalculationType): void;
  }

  enum CalculationType {
    recalculate = "Recalculate"
  }
}

declare namespace OfficeScript {
  interface MailProperties {
    subject: string;
  }
}
