import { IBaseTemplateModel } from "./IBaseTemplateModel";
import { IContentType } from "./IContentType";
import { IField } from "./IField";

export interface IList extends IBaseTemplateModel {
    UniqueId: string;
    Title: string;
    Url: string;
    TemplateType: string;
    Description?: string;
    Hidden?: boolean;
    EnableAttachments?: boolean;
    EnableFolderCreation?: boolean;
    EnableVersioning?: boolean;
    MinorVersionLimit?: number;
    MaxVersionLimit?: number;
    DraftVersionVisibility?: number;

    ContentTypeRefIds?: any[];
    Views?: any[];

    setAvailableFieldsAndContentTypes(fields: IField[], contentTypes: IContentType[]): void;
}