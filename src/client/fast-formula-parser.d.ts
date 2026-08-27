/**
 * fast-formula-parser(MIT)에는 공식 타입 선언이 없다 — README/소스(node_modules/fast-formula-parser)를
 * 직접 읽고 실제 쓰는 부분만 최소한으로 선언한다.
 */
declare module 'fast-formula-parser' {
  export interface CellRef {
    readonly sheet: string;
    readonly row: number;
    readonly col: number;
  }

  export interface RangeRef {
    readonly sheet: string;
    readonly from: { readonly row: number; readonly col: number };
    readonly to: { readonly row: number; readonly col: number };
  }

  export class FormulaError extends Error {
    static readonly DIV0: FormulaError;
    static readonly NA: FormulaError;
    static readonly NAME: FormulaError;
    static readonly NULL: FormulaError;
    static readonly NUM: FormulaError;
    static readonly REF: FormulaError;
    static readonly VALUE: FormulaError;
    readonly name: string; // 예: '#DIV/0!'
  }

  export const Types: {
    readonly NUMBER: unknown;
    readonly STRING: unknown;
    readonly BOOLEAN: unknown;
    readonly ARRAY: unknown;
  };

  export interface ParsedCriteria {
    readonly op: string;
    readonly value: unknown;
    readonly match?: boolean;
  }

  export const Criteria: {
    parse(criteria: unknown): ParsedCriteria;
  };

  export const WildCard: {
    isWildCard(value: unknown): boolean;
    toRegex(value: unknown, flags?: string): RegExp;
  };

  export interface FormulaHelpersApi {
    accept(param: unknown, type?: unknown, defValue?: unknown, flat?: boolean, allowSingleValue?: boolean): any;
    flattenParams(
      params: readonly unknown[],
      valueType: unknown,
      allowUnion: boolean,
      hook: (item: unknown, info: { isLiteral: boolean }) => void,
    ): void;
    retrieveArg(context: FormulaContext, arg: unknown): { value: unknown; isArray: boolean };
  }

  export const FormulaHelpers: FormulaHelpersApi;

  export interface FormulaContext {
    readonly position: { readonly row: number; readonly col: number; readonly sheet: string };
    readonly utils: { extractRefValue(arg: unknown): { val: unknown; isArray: boolean } };
  }

  export type FormulaFunction = (...args: any[]) => unknown;
  export type ContextFormulaFunction = (context: FormulaContext, ...args: any[]) => unknown;

  export interface FormulaParserOptions {
    functions?: Record<string, FormulaFunction>;
    functionsNeedContext?: Record<string, ContextFormulaFunction>;
    onCell?: (ref: CellRef) => unknown;
    onRange?: (ref: RangeRef) => unknown[][];
    onVariable?: (name: string, sheetName?: string) => CellRef | RangeRef;
  }

  export interface FormulaPosition {
    readonly row: number;
    readonly col: number;
    readonly sheet: string;
  }

  export default class FormulaParser {
    constructor(options?: FormulaParserOptions);
    parse(expression: string, position: FormulaPosition, allowReturnArray?: boolean): unknown;
  }
}
