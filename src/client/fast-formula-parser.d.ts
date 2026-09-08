/**
 * fast-formula-parser(MIT)에는 공식 타입 선언이 없다 — README/소스(node_modules/fast-formula-parser)를
 * 직접 읽고 실제 쓰는 부분만 최소한으로 선언한다.
 *
 * 주의: 이 모듈은 CommonJS(`module.exports = FormulaParser` 클래스 + `Object.assign`으로 나중에
 * 덧붙인 정적 프로퍼티들)라서, Node ESM 환경(tsx/실제 서버 프로세스)의 cjs-module-lexer 정적 분석은
 * `Criteria`/`FormulaError`/`FormulaHelpers`/`Types`/`WildCard`를 별도 named export로 인식하지
 * 못한다(vitest는 esbuild/Vite의 런타임 인터롭 셔임을 쓰므로 겉보기엔 통과하지만, 실제 서버는
 * `SyntaxError: does not provide an export named ...`로 죽는다). 그래서 이들은 여기서도 top-level
 * named export가 아니라 default-export된 FormulaParser 클래스의 static 멤버로 선언하고, 사용하는
 * 쪽에서는 `import FormulaParser from 'fast-formula-parser'; const { Criteria, ... } = FormulaParser;`
 * 형태(default import 후 구조분해)로 가져와야 한다.
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

  class FormulaErrorClass extends Error {
    static readonly DIV0: FormulaErrorClass;
    static readonly NA: FormulaErrorClass;
    static readonly NAME: FormulaErrorClass;
    static readonly NULL: FormulaErrorClass;
    static readonly NUM: FormulaErrorClass;
    static readonly REF: FormulaErrorClass;
    static readonly VALUE: FormulaErrorClass;
    readonly name: string; // 예: '#DIV/0!'
  }

  interface TypesApi {
    readonly NUMBER: unknown;
    readonly STRING: unknown;
    readonly BOOLEAN: unknown;
    readonly ARRAY: unknown;
  }

  export interface ParsedCriteria {
    readonly op: string;
    readonly value: unknown;
    readonly match?: boolean;
  }

  interface CriteriaApi {
    parse(criteria: unknown): ParsedCriteria;
  }

  interface WildCardApi {
    isWildCard(value: unknown): boolean;
    toRegex(value: unknown, flags?: string): RegExp;
  }

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

    // Object.assign(FormulaParser, { Criteria, FormulaError, FormulaHelpers, Types, WildCard, ... })로
    // 런타임에 덧붙는 정적 프로퍼티들 — 실제 CJS 구조를 그대로 타입에 반영한다.
    static readonly FormulaError: typeof FormulaErrorClass;
    static readonly Types: TypesApi;
    static readonly Criteria: CriteriaApi;
    static readonly WildCard: WildCardApi;
    static readonly FormulaHelpers: FormulaHelpersApi;
  }
}
