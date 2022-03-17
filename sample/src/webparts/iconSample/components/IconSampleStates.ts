import { IColumn } from "@fluentui/react";

export interface IDocument {
    key: string;
    name: string;
    value: string;
    iconName: string;
    fileType: string;
    modifiedBy: string;
    dateModified: string;
    dateModifiedValue: number;
    fileSize: string;
    fileSizeRaw: number;
}

export interface IIconSampleStates {
    columns: IColumn[];
    items: IDocument[];
    isCompactMode: boolean;
}
