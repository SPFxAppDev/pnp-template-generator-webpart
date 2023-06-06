import { IBaseTemplateModel } from "./IBaseTemplateModel";
import { IField } from "./IField";

export interface ISiteFields extends IBaseTemplateModel {
    fields: IField[];
}