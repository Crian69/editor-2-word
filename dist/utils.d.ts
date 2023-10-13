declare function typeOf(obj: unknown): any;
export { typeOf };
export declare const isFilledArray: (arr: unknown) => boolean;
export declare const getUniqueArrayByKey: <T>(arr: T[], uniqueKey?: string) => T[];
export declare const removeTagDIV: (str: string) => string;
export declare const escape2Html: (str: string) => string;
export declare const trimHtml: (str: string) => string;
export declare const deepCopyByJSON: <T>(obj: T) => T;
export declare const isValidColor: (color: string) => boolean;
export declare const toHex: (color: string) => string;
import { SizeNumber } from './types';
/**
 * parse size
 */
export declare const handleSizeNumber: (val: string) => SizeNumber;
export declare const numberCM: (size: string) => number;
export declare const calcMargin: (margin: string) => number;
export declare const optimizeBlankSpace: (content: string, ratio?: number) => string;
export declare const getImageBlob: (src: string) => Promise<Blob>;
