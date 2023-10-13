import { Paragraph } from 'docx';
interface ListNode {
    type: string;
    children: Array<{
        content: string;
    }>;
}
export declare function buildList(node: ListNode): Paragraph[];
export {};
