import { StyleInterface, StyleOption } from './../types';
export declare const tokens: {
    name: string;
    judge: ({ key }: StyleInterface) => boolean;
    handler: import("./types").TokenHandler;
}[];
export declare const provideStyle: (styles: StyleInterface[]) => StyleOption;
