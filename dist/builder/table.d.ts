import { BorderStyle, ITableOptions, Table } from 'docx';
import { CustomTagStyleMap, Node } from '../types';
export declare const calcTableWidth: (colsArr: number[]) => number;
export declare const getTableBorderStyleSingle: (size: number, color: string) => {
    style: BorderStyle;
    size: number;
    color: string;
};
export declare const getColGroupWidth: (cols: Node[]) => number[];
export declare const handleCellWidthFromColgroup: (cols: number[], index: number, colspan: number) => number;
export declare const getCellWidthInDXA: (size: number) => number;
export declare const tableNodeToITableOptions: (tableNode: Node, tagStyleMap?: CustomTagStyleMap) => Promise<ITableOptions | null>;
export declare const tableCreator: (tableNode: Node, tagStyleMap?: CustomTagStyleMap) => Promise<Table | null>;
