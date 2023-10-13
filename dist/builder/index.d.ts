import { Paragraph } from 'docx';
import { CustomTagStyleMap, Node } from '../types';
export declare const contentBuilder: (node: Node, tagStyleMap?: CustomTagStyleMap) => Promise<Paragraph | import("docx").Table | null>;
