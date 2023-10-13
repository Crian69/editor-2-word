import { CustomTagStyleMap, Node, StyleInterface, StyleOption } from '../types';
import { ParagraphChild, TextRun } from 'docx';
export declare const toFlatStyleList: (styleStringList: string[]) => StyleInterface[];
export declare const calcTextRunStyle: (styleList: string[], tagStyleMap?: CustomTagStyleMap) => Partial<StyleOption>;
export declare const textCreator: (node: Node, tagStyleMap?: CustomTagStyleMap) => TextRun;
export declare const getChildrenByTextRun: (nodeList: Node[], tagStyleMap?: CustomTagStyleMap) => Promise<ParagraphChild[]>;
