import { AlignmentType, BorderStyle, VerticalAlign } from 'docx';
import { IPageLayout } from './types';
export declare const Splitter_Colon = ":";
export declare const Splitter_Semicolon = ";";
export declare const PXbyTWIPS = 16;
export declare const PXbyPT: number;
export declare const D_FontSizePX = 16.3;
export declare const D_FontSizePT: number;
export declare const D_LineHeight = 1.5;
export declare const D_PageWidthPX = 794;
export declare const D_PageHeightPX = 1123;
export declare const D_PagePaddingPX = 71;
export declare const D_PageTableFullWidth = 642;
export declare const D_TableFullWidth = 9035;
export declare const D_TableBorderColor = "444444";
export declare const A4MillimetersWidth = 145.4;
export declare const D_CELL_MARGIN: number;
export declare const D_TableBorderSize = 2;
export declare const D_TableCellHeightPx = 18;
export declare const FontSongTi: string[];
export declare const AlignMap: {
    left: AlignmentType;
    center: AlignmentType;
    right: AlignmentType;
};
export declare const hyperlinkColor = "#007AFF";
export declare const D_TagStyleMap: {
    p: string;
    strong: string;
    em: string;
    u: string;
    del: string;
    h1: string;
    h2: string;
    h3: string;
    h4: string;
    h5: string;
    h6: string;
    sub: string;
    sup: string;
    a: string;
};
export declare const D_Layout: IPageLayout;
export declare const Direction: {
    left: string;
    right: string;
    firstLine: string;
    start: string;
    end: string;
    hanging: string;
};
export declare const PaddingDirection: {
    'padding-left': string;
    'padding-right': string;
    'padding-top': string;
    'padding-bottom': string;
};
export declare const Size: {
    em: string;
    px: string;
    pt: string;
};
export declare const SingleLine: {
    type: string;
    color: string;
};
export declare const TagType: {
    table: string;
    link: string;
    text: string;
    img: string;
    ordered_list: string;
    unordered_list: string;
};
export declare const DefaultBorder: {
    style: BorderStyle;
    size: number;
    color: string;
};
export declare const verticalAlignMap: {
    top: VerticalAlign;
    middle: VerticalAlign;
    bottom: VerticalAlign;
};
