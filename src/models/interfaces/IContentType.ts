import { IBaseTemplateModel } from "./IBaseTemplateModel";
import { IField } from "./IField";

export interface IContentType extends IBaseTemplateModel {
    ID: string;
    ParentID: string;
    Name: string;
    Description?: string;
    Group?: string;
    Hidden?: boolean;
    FieldRefIds?: string[];
    getInternalIdentifier(): string;
    setAllAvailableCustomSiteColumns(fields: IField[]): void;
}